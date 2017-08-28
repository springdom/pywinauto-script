import time
import pywinauto
from pywinauto import application
from pywinauto.keyboard import SendKeys
from pywinauto import Desktop
from pywinauto.application import Application
from pywinauto.findwindows import find_window
from openpyxl import load_workbook
from string import ascii_lowercase

app = Application(backend='uia')
p = pywinauto.findwindows.find_element(best_match="Interaction Administrator")
app.connect(handle=p.handle)
dlg = app.window(best_match="Interaction Administrator")
typein = app.dlg.type_keys

#app.dlg.print_control_identifiers() #Check Identifiers

wb = load_workbook('orgchart.xlsx', read_only=True)
sh = wb['HGV_OrgChart']
ws = wb.active

orl_outbnd_cms = ["MKT-Outbound-Callback", "MKT-Outbound-Main2", "Orl_OUT_SUP"]
orl_outbnd_sf = ["LOC-ORL-MKT-SalesForce", "SF-Orlando-Manual", "SF-RestrictDialing"] #Manual
orl_ct = ["CT Priority 1", "CT Priority 2", "LOC-ORL-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
orl_act = ["LOC-ORL-MKT-ACT", "MKT-Activations-CallBack", "MKT-ACT-Main", "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]
orl_cc = ["LOC-ORL-MKT-CC",  "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]

spg_outbnd_cms = ["MKT-Outbound-Callback", "MKT-Outbound-Main2", "SPG_OUT_SUP"]
spg_outbnd_sf = ["LOC-SPG-MKT-SalesForce", "SF-Springfield-Manual", "SF-RestrictDialing"]
spg_ct = ["LOC-SPG-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]

lv_outbnd_cms = ["MKT-Outbound-Callback", "MKT-Outbound-Main2", "LV_OUT_SUP"]
lv_outbnd_sf = ["LOC-LV-MKT-SalesForce", "SF-LV-Manual", "SF-RestrictDialing"]
lv_ct = ["CT Priority 1", "CT Priority 2", "LOC-LV-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]

licenses = ["Interaction Optimizer Client Access","Interaction Optimizer Real-time Adherence Tracking","Interaction Optimizer Schedulable"]
roles = ["MKT-Agent","MKT-SF-Agent","MKT-CC-Agent"]

column_header = {}
Add = "Add/Delete/Change/Transfer/Rehire"
Add2 = "Add/Delete/Change"
windows = "Windows for Adds"
email = "Email Address"
Name = "AgentName"
cic_id = "CIC_ID"

def main():
    get_alphabet()
    column_headers()

def get_alphabet():
    for c in ascii_lowercase:
	x = sh[c.upper() + "1"].value
	column_header[c.upper()] = sh[c.upper() + "1"].value

def sed(a):
     for k,v in column_header.items():
         if v == a:
             a = k
         return str(k)

def column_headers():
    add = sed(Add) or sed(Add2)
    agent_username = sed(windows)
    agentemail = sed(email) or "AZ"        
    agent_name =  sed(Name)
    agent_tsr = sed(cic_id)

    orgchart_data(add, agent_username, agentemail,agent_name, agent_tsr)

def orgchart_data(add, windows, agent_email,agent_name, agent_tsr):
    n = 2
    while n < sh.max_row :
        if sh[add + str(n)].value == "Add":
            username = sh[windows + str(n)].value
            agentName = sh[agent_name + str(n)].value
            email = sh[agent_email + str(n)].value
            tsr = sh[agent_tsr + str(n)].value
            
            print(username, email,agentName, tsr)
            
            Config(tsr, username)
            GetUserDetails(agentName)
            AutoACD()
            Roles(department)
            getWorkGroups()
            if department == "ct" or department == "cc":
                Licensing()
            #Email
            """
            if loc == "orl":
                Email(email)
            """
            app.dlg.Cancel.click_input() #Change When DOne
        n += 1

def getWorkGroups(): # helper function here
    if location == "orl":
        if department == "outbnd":
            AgentWorkGroups(orl_outbnd_cms)
        if department == "ct":
            AgentWorkGroups(orl_ct)
        if department == "act":
            AgentWorkGroups(orl_act)
        if department == "cc":
            AgentWorkGroups(orl_cc)

    if location == "spg":
        if department == "outbnd":
            AgentWorkGroups(spg_outbnd_cms)
        if department == "ct":
            AgentWorkGroups(spg_ct)

    if location == "lvn":
        if department == "outbnd":
            AgentWorkGroups(lv_outbnd_cms)
        if department == "ct":
            AgentWorkGroups(lv_ct)

def Config(tsr, win_username):
    app.dlg.menu_select("File->New")
    app.dlg.Edit0.type_keys(tsr + "{ENTER}")

    app.dlg.Edit3.type_keys("10102015") #Password
    app.dlg.Edit4.type_keys("10102015") #Confirm Password
    app.dlg.Edit7.type_keys("hgvcnt\\" + win_username) #Domain User
    #Location?

#Personal Info
def GetUserDetails(agentName):
    name = agentName.split(" ")

    app.dlg.TabItem3.click_input()
    app.dlg.Edit0.type_keys(name[0]) #FirstName
    app.dlg.Edit2.type_keys(name[1]) #LastName
    app.dlg.Edit4.type_keys(name[0] + "{SPACE}" + name[1]) #Display Name

def AgentWorkGroups(wrkgrps):
    num = 0
    app.dlg.Workgroups.click_input()
    app.dlg.OK.click_input()

    for x in wrkgrps:
        if department == "outbnd":
            if num == 0:
                app.dlg['ListBox'].click_input()
                app.dlg['ListBox'].wheel_mouse_input(wheel_dist = -13)
            if num == 2:
                app.dlg['ListBox'].click_input()
                app.dlg['ListBox'].wheel_mouse_input(wheel_dist = -5)

        num+=1
        app.dlg[x].click_input()
        app.dlg.Add.click_input()

        if department == "ct":
            if num == 3:
                app.dlg['ListBox'].click_input()
                app.dlg['ListBox'].wheel_mouse_input(wheel_dist = -10)
        if department == "act" or department == "cc":
            if num == 1:
                app.dlg['ListBox'].click_input()
                app.dlg['ListBox'].wheel_mouse_input(wheel_dist = -6)

def AutoACD():
    app.dlg.ACD.click_input()
    app.dlg.ListItem3.click_input()
    app.dlg.CheckBox0.click_input() #Auto ACD

def Roles(deptmnt):
    app.dlg.Roles.click_input()
    app.dlg.Add.click_input()

    if location == "orl":
        if deptmnt == "ct" or deptmnt == "outbnd":
            app.dlg.ListItem4.select()
        elif deptmnt == "act" or deptmnt == "cc":
            app.dlg.ListItem6.select()
    if location == "spg":
        pass
    if location == "lvn":
        pass

    app.dlg.OK.click_input()

def Licensing():
    app.dlg.Licensing.click_input()
    app.dlg.OK.click_input()
    app.dlg['Enable Licenses'].click_input()
    for ls in licenses:
        app.dlg[ls].type_keys("{SPACE}")

def Email(email):
    app.dlg['Configuration'].click_input()
    app.dlg['...'].click_input()
    app.dlg['RadioButton6'].click_input()
    app.dlg['Email address'].click_input()
    app.dlg["Edit"].type_keys("{BS 40}")
    app.dlg["Edit"].type_keys(email)

location = input("Location orl, spg, lvn: ")
department = input("Department outbnd, ct, act, cc: ")

main()


"""
Agent Queues
---------------------------------------------------------------------
No Licenses
Roles - MKT-Agent
- Orlando - OutBoundCMS
Workgroups - MKT-Outbound-Callback, MKT-Outbound-Main2, Orl_OUT_SUP

- Las Vegas - OutBoundCMS
Workgroups - MKT-Outbound-Callback, MKT-Outbound-Main2, LV_OUT_SUP

- SpringField - OutBoundCMS
Workgroups - MKT-Outbound-Callback, MKT-Outbound-Main2, SPG_OUT_SUP

---------------------------------------------------------------------
No Licenses
Roles - MKT-SF-Agent
- Orlando - OutBoundManual
LOC-ORL-MKT-SalesForce, SF-Orlando-Manual, SF-RestrictDialing

- Las Vegas - OutBoundManual
LOC-LV-MKT-SalesForce, SF-LV-Manual, SF-RestrictDialing

- SpringField - OutBoundManual
LOC-SPG-MKT-SalesForce, SF-Springfield-Manual, SF-RestrictDialing

---------------------------------------------------------------------
Enable Licenses - Interaction Optimizer CLient Access,
Interaction Optimizer Real Time Adherance Tracking,
Interaction Optimizer Scheduable
Roles - MKT-Agent
- Orlando - Call Tranfer
Workgroups - CT Priority 1, CT Priority 2, LOC-ORL-MKT-HRCC, MKT-InbCT-Callback, MKT-InbCT-HRCC

Las Vegas - Call Transfer
CT Priority 1, CT Priority 2, LOC-LV-MKT-HRCC, MKT-InbCT-Callback, MKT-InbCT-HRCC

- SpringField - Call Transfer
Workgroups - LOC-SPG-MKT-HRCC, MKT-InbCT-Callback, MKT-InbCT-HRCC

Call Transfer - Client Optimizer
---------------------------------------------------------------------
- Orlando Activations
No Licenses
WorkGroups - LOC-ORL-MKT-ACT, MKT-Activations-CallBack, MKT-ACT-Main, MKT-CC-BookDates, MKT-CC-BookDates-Priority1,MKT-CC-CustomerService, MKT-CC-CustomerService-Priority2
Roles - MKT-CC-Agent
---------------------------------------------------------------------
- ORlando Customer Care
Enable Licenses - Interaction Optimizer CLient Access, Interaction Optimizer Real Time Adherance Tracking, Interaction Optimizer Scheduable
WorkGroups - LOC-ORL-MKT-CC,  MKT-CC-BookDates, MKT-CC-BookDates-Priority1,MKT-CC-CustomerService, MKT-CC-CustomerService-Priority2
Roles - MKT-CC-Agent
---------------------------------------------------------------------
Auto-ACD
"""
