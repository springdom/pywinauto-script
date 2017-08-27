from openpyxl import load_workbook
from string import ascii_lowercase

wb = load_workbook('orgchart.xlsx', data_only=True)
sh = wb['HGV_OrgChart']
ws = wb.active
column_header = {}
#Used Columns
Add = "Add/Delete/Change/Transfer/Rehire"
Add2 = "Add/Delete/Change"
windows = "Windows for Adds"
email = "Email Address"
Name = "AgentName"
cic_id = "CIC_ID"



#loop through and find column name
for c in ascii_lowercase:
	x = sh[c.upper() + "1"].value
	column_header[c.upper()] = sh[c.upper() + "1"].value

def orgchart_data(add, windows, agent_name, agent_tsr):
    n = 2
    while sh[add + str(n)].value != None or n < 180:
        if sh[add + str(n)].value == "Add":
            username = sh[windows + str(n)].value
            agentName = sh[agent_name + str(n)].value
            tsr = sh[agent_tsr + str(n)].value
            
            print(username, agentName, tsr)
        n += 1

def cheese():
    for k,v in column_header.items():
        if v == Add:
            add = k
        if v == windows:
            agent_username = k
        """
        if v == email:
            agent_email = k
        """
        if v == Name:
            agent_name = k
        if v == cic_id:
            agent_tsr = k
            
    orgchart_data(add, agent_username, agent_name, agent_tsr)
            
cheese()
 

