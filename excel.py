from openpyxl import load_workbook

wb = load_workbook('KeyCodes.xlsx', data_only=True)
sh=wb["Sheet1"]
ws = wb.active

outbnd_orl = {}
orl_call_trsf = {}
lvn_call_trsf = {}
gldmtn_call_trsf = {}

#KeyCodes - Outbound, Orlando, Las Vegas, Gold Mountain
def keycodes(m_row,mx_col,mx_row,x):
    num = 0
    for row in ws.iter_rows(min_row=m_row, max_col=mx_col, max_row=mx_row):
        for cell in row:
            num += 1
            if num == 1:
                #print(cell.value)
                x[cell.value] = None
            if m_row > 17:
                if num == 33:
                    num = 0
            else:
                if num == 3:
                    num = 0

def assign_key(n,x):
    if n > 17:
        for k,v in x.items():
            x[k] = sh['D%s' % n].value
            n += 1
    else:
        for k,v in x.items():
            x[k] = sh['B%s' % n].value
            n += 1

def tes(z):
    for k in z.items():
        print(k)
        """
    mydict = {
        'orl':,
        'nyc':,
        'lvn':,
        'hhv':,
        'haw':,
        'cal':,
        'sca':,
        'mtn':,
        'wdc':,
        }
        """

keycodes(4,3,12,outbnd_orl)
assign_key(4,outbnd_orl)
    
keycodes(18,33,27,orl_call_trsf)
assign_key(18,orl_call_trsf)

keycodes(32,33,41,lvn_call_trsf)
assign_key(32,lvn_call_trsf)

keycodes(46,33,55,gldmtn_call_trsf)
assign_key(46,gldmtn_call_trsf)
