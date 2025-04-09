from flask import Flask
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re
from collections import defaultdict
from ortools.linear_solver import pywraplp

app = Flask(__name__)

def run_solver():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('prototren-dda15b5a1ed4.json', scope)
        client = gspread.authorize(creds)

        sheet_formulasi = client.open("Prototype5").worksheet("Formulasi")
        objective_raw = sheet_formulasi.acell("C4").value.replace(" ", "")
        matches = re.findall(r'([+-]?\d*\.?\d*)X(\d+)', objective_raw)

        coeff_dict = defaultdict(float)
        for coef_str, var_index in matches:
            if coef_str in ["", "+"]:
                coef = 1.0
            elif coef_str == "-":
                coef = -1.0
            else:
                coef = float(coef_str)
            coeff_dict[int(var_index)] += coef

        n_vars = max(coeff_dict.keys())
        coeffs = [coeff_dict[i+1] for i in range(n_vars)]

        constraints_raw = sheet_formulasi.col_values(2)[23:38]
        operators = sheet_formulasi.col_values(5)[5:20]
        rhs_values = sheet_formulasi.col_values(6)[5:20]

        solver = pywraplp.Solver.CreateSolver("SCIP")
        variables = [solver.IntVar(0, 1, f"x{i+1}") for i in range(len(coeffs))]
        solver.Maximize(solver.Sum([coeffs[i] * variables[i] for i in range(len(coeffs))]))

        parsed_constraints = []
        for i in range(len(constraints_raw)):
            constraint = constraints_raw[i].replace(" ", "")
            operator = operators[i]
            rhs = float(rhs_values[i])
            terms = re.findall(r'([+-]?\d*\.?\d*)X(\d+)', constraint)
            expr = []
            term_list = []
            for coef_str, var_index in terms:
                if coef_str in ["", "+"]:
                    coef = 1.0
                elif coef_str == "-":
                    coef = -1.0
                else:
                    coef = float(coef_str)
                idx = int(var_index) - 1
                expr.append(coef * variables[idx])
                term_list.append((coef, idx))
            parsed_constraints.append((term_list, operator, rhs))
            if operator == "<=":
                solver.Add(solver.Sum(expr) <= rhs)
            elif operator == "=":
                solver.Add(solver.Sum(expr) == rhs)
            elif operator == ">=":
                solver.Add(solver.Sum(expr) >= rhs)

        status = solver.Solve()
        sheet_output = client.open("Prototype5").worksheet("Output")
        sheet_output.batch_clear(['A1:C1000'])

        if status == pywraplp.Solver.OPTIMAL:
            hasil = [int(var.solution_value()) for var in variables]
            nilai_z = solver.Objective().Value()

            slack_list = []
            for terms, operator, rhs in parsed_constraints:
                total = sum(coef * hasil[idx] for coef, idx in terms)
                if operator == "<=":
                    slack = rhs - total
                elif operator == ">=":
                    slack = total - rhs
                else:
                    slack = 0
                slack_list.append(slack)

            n_sol = len(hasil)
            n_slack = len(slack_list)
            n_rows = max(n_sol, n_slack) + 1
            table = [["" for _ in range(3)] for _ in range(n_rows)]
            table[0] = ["Solusi", "Nilai Z", "Slack/Surplus"]
            for i in range(n_sol):
                table[i+1][0] = hasil[i]
            table[1][1] = nilai_z
            for i in range(n_slack):
                table[i+1][2] = slack_list[i]

            sheet_output.update("A1", table)
            return f"✅ Solver selesai! Nilai Z = {nilai_z}"
        else:
            sheet_output.update("A1", [["❌ Tidak ditemukan solusi optimal."]])
            return "❌ Tidak ditemukan solusi optimal."

    except Exception as e:
        return f"❌ Error: {str(e)}"

@app.route('/')
def home():
    return '🟢 Server aktif! Kunjungi /trigger untuk mulai solver.'

@app.route('/trigger')
def trigger_solver():
    hasil = run_solver()
    return hasil

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=3000)