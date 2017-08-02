from openpyxl import load_workbook

wb = load_workbook('KeyCodes.xlsx', data_only=True)
sh=wb["Sheet1"]
ws = wb.active
num = 0
test = []
loc_dict = {}

#KeyCodes - Outbound, Orlando, Las Vegas, Gold Mountain

for row in ws.iter_rows(min_row=18, max_col=33, max_row=27):
    for cell in row:
        num += 1
        if num == 1:
            #print(cell.value)
            loc_dict[cell.value] = None
        if num == 33:
            num = 0
            #print(cell.value)
            
