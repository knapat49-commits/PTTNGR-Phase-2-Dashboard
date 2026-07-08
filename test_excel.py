import os, openpyxl

path = os.path.join(os.path.dirname(__file__),
                    "Weekly Progress Phase2 update on 2026-06-17.xlsx")

wb = openpyxl.load_workbook(path, data_only=True)
data = wb["DATA"]
chart = wb["CHART"]

print("=== DATA sheet S-curve rows (rows 4-6, cols H onwards) ===")
# Row 4=TOR, 5=mplan, 6=actual; col 8=H (wk0), 9=I (wk1), ...
for row_num, label in [(4,"TOR"), (5,"mplan"), (6,"actual")]:
    vals = []
    for col in range(8, 25):  # cols H through X
        v = data.cell(row_num, col).value
        vals.append(v)
    print(f"  {label}: {vals}")

print("\n=== CHART sheet progress rows (cols G=plan, H=actual) ===")
for row in [7,8,9,10,11,12,15,16,17,18,19,20,23,24,25,26,27,28]:
    g = chart.cell(row, 7).value
    h = chart.cell(row, 8).value
    print(f"  row {row:2d}: plan={g}  actual={h}")

print("\n=== Bar chart data (rows 8-12, col H) ===")
for row in range(8, 13):
    v = data.cell(row, 8).value
    print(f"  row {row}: {v}")
