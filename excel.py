from openpyxl import load_workbook

wb = load_workbook('KeyCodes.xlsx', data_only=True)
sh=wb["Sheet1"]
ws = wb.active
test = []
for row in ws.iter_rows(min_row=18, max_col=33, max_row=18):
    for cell in row:
        test.append(cell.value)
        print(cell.value)
