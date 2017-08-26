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
#LAS,ORL,SPG
#CT,ACT,CC,Outbound

wb = load_workbook('orgchart.xlsx', data_only=True)
sh = wb['HGV_OrgChart']
ws = wb.active

orl_outbnd_cms = ["MKT-Outbound-Callback", "MKT-Outbound-Main2", "Orl_OUT_SUP"]
orl_outbnd_sf = ["LOC-ORL-MKT-SalesForce", "SF-Orlando-Manual", "SF-RestrictDialing"] #Manual
orl_ct = ["CT Priority 1", "CT Priority 2", "LOC-ORL-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
orl_act = ["LOC-ORL-MKT-ACT", "MKT-Activations-CallBack", "MKT-ACT-Main", "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]
orl_cc = ["LOC-ORL-MKT-CC",  "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]

licenses = ["Interaction Optimizer Client Access","Interaction Optimizer Real-time Adherence Tracking","Interaction Optimizer Schedulable"]
roles = ["MKT-Agent","MKT-SF-Agent","MKT-CC-Agent"]

def main():
    pass

"""
loop through and find column name
for c in ascii_lowercase:
	print(c)
"""

def orgchart_data():
    n = 2
#Find Correct columns ?Give option to setup correct colums then check if correct column
    for x in range(1,200):
        if sh['A%s' % n].value == "Add":
            username = sh['B%s' % n].value
            agentName = sh['D%s' % n].value
            tsr = str(sh['E%s' % n].value)
            #Email
            
            print(username, agentName, tsr)
            
            Config(tsr, username)
            GetUserDetails(agentName)
            AutoACD()
            Roles(department)
            getWorkGroups()
            if department == "ct" or department == "cc":
                Licensing("Test")
            else:
                pass
            app.dlg.Cancel.click_input() #Change When DOne
            n += 1

def getWorkGroups():
    if location == "orl":
        if department == "ct":
            AgentWorkGroups(orl_ct)
        if department == "act":
            AgentWorkGroups(orl_act)
        if department == "cc":
            AgentWorkGroups(orl_cc)
    elif location == "spg":
        pass
    elif location == "lvn":
        pass

def Config(tsr, win_username):
    app.dlg.menu_select("File->New")
    app.dlg.Edit0.type_keys(tsr + "{ENTER}")

    #app.dlg.Edit.type_keys("Matthew") #Extension
    app.dlg.Edit3.type_keys("10102015") #Password
    app.dlg.Edit4.type_keys("10102015") #Confirm Password
    app.dlg.Edit7.type_keys("hgvcnt\\" + win_username) #Domain User
    #Email
    #Location?


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
        num+=1
        if department == "ct":
            app.dlg[x].click_input()
            app.dlg.Add.click_input()
            if num == 3:
                app.dlg['CC Elite 1'].click_input()
                app.dlg['ListBox'].wheel_mouse_input(wheel_dist = -10)
        if department == "act" or department == "cc":
            app.dlg[x].click_input()
            app.dlg.Add.click_input()
            if num == 1:
                app.dlg['CC Elite 1'].click_input()
                app.dlg['ListBox'].wheel_mouse_input(wheel_dist = -6)

def QueueScroll():
    pass
                              
def AutoACD():
    app.dlg.ACD.click_input()
    app.dlg.ListItem3.click_input()
    app.dlg.CheckBox0.click_input() #Auto ACD

def Roles(deptmnt):
    app.dlg.Roles.click_input()
    app.dlg.Add.click_input()
    
    if deptmnt == "ct":
        app.dlg.ListItem4.select()
    elif deptmnt == "act" or deptmnt == "cc":
        app.dlg.ListItem6.select()

    app.dlg.OK.click_input()

def Licensing(department):
    app.dlg.Licensing.click_input()
    app.dlg.OK.click_input()
    app.dlg['Enable Licenses'].click_input()
    for ls in licenses:
        app.dlg[ls].type_keys("{SPACE}")

    """"
    app.dlg['Interaction Optimizer Client Access'].type_keys("{SPACE}")
    app.dlg['Interaction Optimizer Real-time Adherence Tracking'].type_keys("{SPACE}")
    app.dlg['Interaction Optimizer Schedulable'].type_keys("{SPACE}")
    """

location = input("Location orl, spg, lvn: ")
department = input("Department ct, act, cc: ")

orgchart_data()


"""
Agent Queues
---------------------------------------------------------------------
- Orlando - OutBoundCMS
No Licenses
Workgroups - MKT-Outbound-Callback, MKT-Outbound-Main2, Orl_OUT_SUP
Roles - MKT-Agent
- Las Vegas - OutBoundCMS
No Licenses
Workgroups - MKT-Outbound-Callback, MKT-Outbound-Main2, LV_OUT_SUP
Roles - MKT-Agent
- SpringField - OutBoundCMS
No Licenses
Workgroups - MKT-Outbound-Callback, MKT-Outbound-Main2, SPG_OUT_SUP
Roles - MKT-Agent
---------------------------------------------------------------------
- Orlando - OutBoundManual
No Licenses
LOC-ORL-MKT-SalesForce, SF-Orlando-Manual, SF-RestrictDialing
Roles - MKT-SF-Agent
- Las Vegas - OutBoundManual
No Licenses
LOC-LV-MKT-SalesForce, SF-LV-Manual, SF-RestrictDialing
Roles - MKT-SF-Agent
- SpringField - OutBoundManual
No Licenses
LOC-SPG-MKT-SalesForce, SF-Springfield-Manual, SF-RestrictDialing
Roles - MKT-SF-Agent
---------------------------------------------------------------------
- Orlando - Call Tranfer
Enable Licenses - Interaction Optimizer CLient Access, Interaction Optimizer Real Time Adherance Tracking, Interaction Optimizer Scheduable
Workgroups - CT Priority 1, CT Priority 2, LOC-ORL-MKT-HRCC, MKT-InbCT-Callback, MKT-InbCT-HRCC
Roles - MKT-Agent
- Las Vegas - Call Transfer
Enable Licenses - Interaction Optimizer CLient Access, Interaction Optimizer Real Time Adherance Tracking, Interaction Optimizer Scheduable
Workgroups - CT Priority 1, CT Priority 2, LOC-LV-MKT-HRCC, MKT-InbCT-Callback, MKT-InbCT-HRCC
Roles - MKT-Agent
- SpringField - Call Transfer
Enable Licenses - Interaction Optimizer CLient Access, Interaction Optimizer Real Time Adherance Tracking, Interaction Optimizer Scheduable
Workgroups - LOC-SPG-MKT-HRCC, MKT-InbCT-Callback, MKT-InbCT-HRCC
Roles - MKT-Agent
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

