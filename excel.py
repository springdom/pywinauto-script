from openpyxl import load_workbook

filed = r"Z:\Packages\Breakdown Of Current Packages 08-04-2017.xlsx"
wb = load_workbook(filed, data_only=True)
sh=wb["Current Packages"]
ws = wb.active

orl_pkgs = {}
lvn_pkgs = {}
nyc_pkgs = {}
haw_pkgs = {}
hhv_pkgs = {}
cal_pkgs = {}
sca_pkgs = {}
mtn_pkgs = {}
wdc_pkgs = {}
listn = []

#KeyCodes - Outbound, Orlando, Las Vegas, Gold Mountain
def keycodes(m_row,mx_col,mx_row,x):
    num = 0
    for row in ws.iter_rows(min_row=m_row, max_col=mx_col, max_row=mx_row):
        for cell in row:
            num += 1
            if num == 7:
                num = 1
            print(num)
            #if cell.value != None:
             #   print(cell.value)
            #print(cell.value)
            #x[cell.value] = None

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

keycodes(4,6,82,orl_pkgs)
#assign_key(4,outbnd_orl)
