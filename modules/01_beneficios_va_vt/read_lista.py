import openpyxl

wb = openpyxl.load_workbook('Conferência Good Card_Base.xlsx', read_only=True, data_only=True)
ws = wb['Lista']
for row in ws.iter_rows(values_only=True):
    non_empty = [str(c) for c in row if c is not None and str(c).strip() != '']
    if non_empty:
        print(" | ".join(non_empty))
