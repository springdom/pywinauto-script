from openpyxl import load_workbook

wb = load_workbook('KeyCodes.xlsx', data_only=True)
sh=wb["Sheet1"]
ws = wb.active

outbnd = {}
orl_call_trsf = {}
lvn_call_trsf = {}
gldmtn_call_trsf = {}



def keycodes(m_row,mx_col,mx_row):
    num = 0
    #KeyCodes - Outbound, Orlando, Las Vegas, Gold Mountain
    for row in ws.iter_rows(min_row=m_row, max_col=mx_col, max_row=mx_row):
        for cell in row:
            num += 1
            if num == 1:
                #print(cell.value)
                orl_call_trsf[cell.value] = None
            if num == 33:
                num = 0

def assign_key(x):
    n = 18
    for k,v in x.items():
        x[k] = sh['D%s' % n].value
        n += 1

keycodes(18,33,27)
assign_key(orl_call_trsf)
print(orl_call_trsf)
