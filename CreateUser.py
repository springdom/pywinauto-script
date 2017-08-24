import time
import pywinauto
from pywinauto import application
from pywinauto.keyboard import SendKeys
from pywinauto import Desktop
from pywinauto.application import Application
from pywinauto.findwindows import find_window

app = Application(backend='uia')
p = pywinauto.findwindows.find_element(best_match="Interaction Administrator")
app.connect(handle=p.handle)
dlg = app.window(best_match="Interaction Administrator")
typein = app.dlg.type_keys
#app.dlg.print_control_identifiers() #Check Identifiers

#LAS,ORL,SPG
#CT,ACT,CC,Outbound

locations = ["orl","llvn","spg"]
departments = ["ct","act","cc","outbound"]

orl_outbnd_cms = ["MKT-Outbound-Callback", "MKT-Outbound-Main2", "Orl_OUT_SUP"]
orl_outbnd_sf = ["LOC-ORL-MKT-SalesForce", "SF-Orlando-Manual", "SF-RestrictDialing"] #Manual
orl_ct = ["CT Priority 1", "CT Priority 2", "LOC-ORL-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
orl_act = ["LOC-ORL-MKT-ACT", "MKT-Activations-CallBack", "MKT-ACT-Main", "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]
orl_cc = ["LOC-ORL-MKT-CC",  "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]

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


def main():
    pass

def NewAgent(): #maybe
    app.dlg.menu_select("File->New")
    app.dlg.Edit0.type_keys("123456" + "{ENTER}")
    
def Config():
    #app.dlg.Edit.type_keys("Matthew") #Extension
    app.dlg.Edit3.type_keys("10102015") #Password
    app.dlg.Edit4.type_keys("10102015") #Confirm Password
    app.dlg.Edit7.type_keys("hgvcnt\\") #Domain User
    #Email
    #Location?
    
def GetUserDetails(name, tsr, username):
    pass

def AgentWorkGroups(loc,deptmnt):
    num = 0
    app.dlg.Workgroups.click_input()
    app.dlg.OK.click_input()
    for x in deptmnt:
        num+=1
        app.dlg[x].click_input()
        app.dlg.Add.click_input()
        if num == 3:
            for x in range(1,4):
                app.dlg['VerticalScrollBar'].click_input()
            
        

def Licensing(licences):
    pass

def AutoACD():
    app.dlg.ACD.click_input()
    app.dlg.ListItem3.click_input()
    app.dlg.CheckBox0.click_input() #Auto ACD

def Roles(dept):
    pass




"""
Check Location
Check Dept
Check Type
"""

"""
TSR
"""
app.dlg.menu_select("File->New")
app.dlg.Edit0.type_keys("123456" + "{ENTER}")

#Configuration
"""
Extension
PassWord
Confirm Password
NT Domain User
"""
#app.dlg.Edit.type_keys("Matthew") #Extension
app.dlg.Edit3.type_keys("10102015") #Password
app.dlg.Edit4.type_keys("10102015") #Confirm Password
app.dlg.Edit7.type_keys("hgvcnt\\") #Domain User

#Personal Info
"""
Last Name
First Name
Display Name
"""
app.dlg.TabItem3.click_input()
app.dlg.Edit0.type_keys("Dave") #FirstName
app.dlg.Edit2.type_keys("Sherman") #LastName
app.dlg.Edit4.type_keys("Dave{SPACE}Sherman") #Display Name

#ACD
"""
Auto ACD
"""
app.dlg.ACD.click_input()
app.dlg.ListItem3.click_input() 
app.dlg.CheckBox0.click_input() #Auto ACD

#Roles
"""
Dept
"""
app.dlg.Roles.click_input()
app.dlg.Add.click_input()
app.dlg.ListItem4.select()
app.dlg.OK.click_input() #Multiple Select?

#WorkGroup
"""
Check Dept
"""

"""
app.dlg.Workgroups.click_input()
app.dlg.OK.click_input()
app.dlg['CT Priority 1'].click_input()
app.dlg.Add.click_input()
app.dlg['CT Priority 2'].click_input()
app.dlg.Add.click_input()
"""

AgentWorkGroups("orl",orl_ct)

#Licensing
app.dlg.Licensing.click_input()
app.dlg.OK.click_input()
app.dlg['Enable Licenses'].click_input()
app.dlg['Interaction Optimizer Client Access'].type_keys("{SPACE}")
app.dlg['Interaction Optimizer Real-time Adherence Tracking'].type_keys("{SPACE}")
app.dlg['Interaction Optimizer Schedulable'].type_keys("{SPACE}")
#app.dlg.OK.click_input()

#pass in obj eg. Edit0
