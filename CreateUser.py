"""
Automates Interaction Administrator
"""
<<<<<<< HEAD
import time
=======
#import time
from string import ascii_lowercase
>>>>>>> parent of b8787b2... formatt
import pywinauto
from pywinauto import application
#from pywinauto.keyboard import SendKeys
#from pywinauto import Desktop
from pywinauto.application import Application
#from pywinauto.findwindows import find_window
from openpyxl import load_workbook
from string import ascii_lowercase
<<<<<<< HEAD

<<<<<<< HEAD
app = Application(backend='uia')
=======
>>>>>>> parent of 55d8f87... a

#CMS
<<<<<<< HEAD
orl_outbnd_cms = ["MKT-Outbound-Callback", "MKT-Outbound-Main2", "Orl_OUT_SUP"]
orl_ct = ["CT Priority 1", "CT Priority 2", "LOC-ORL-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
orl_act = ["LOC-ORL-MKT-ACT", "MKT-Activations-CallBack", "MKT-ACT-Main", "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]
orl_cc = ["LOC-ORL-MKT-CC", "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]
spg_ct = ["LOC-SPG-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
lvn_outbnd_cms = ["LAS_OUT_SUP","MKT-Outbound-Callback", "MKT-Outbound-Main2"]
lv_ct = ["CT Priority 1", "CT Priority 2", "LOC-LV-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]

<<<<<<< HEAD
wrkqueues = {
    "orl_outbnd_cms":orl_outbnd_cms, "orl_ct":orl_ct, "orl_act":orl_act, "orl_cc":orl_cc,
    "spg_ct":spg_ct, "lvn_outbnd_cms":lvn_outbnd_cms, "lvn_ct":lvn_ct,
    }
=======
#CMS
orl_outbnd_cms = ["MKT-Outbound-Callback", "MKT-Outbound-Main2", "Orl_OUT_SUP"]
orl_ct = ["CT Priority 1", "CT Priority 2", "LOC-ORL-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
orl_act = ["LOC-ORL-MKT-ACT", "MKT-Activations-CallBack", "MKT-ACT-Main", "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]
orl_cc = ["LOC-ORL-MKT-CC", "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]
spg_ct = ["LOC-SPG-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
lvn_outbnd_cms = ["LAS_OUT_SUP","MKT-Outbound-Callback", "MKT-Outbound-Main2"]
lv_ct = ["CT Priority 1", "CT Priority 2", "LOC-LV-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]

wrkqueues = {"orl_outbnd_cms":orl_outbnd_cms,"orl_ct":orl_ct,"orl_act":orl_act,"orl_cc":orl_cc,"spg_ct":spg_ct,"lv_outbnd_cms":lv_outbnd_cms,"lv_ct":lv_ct}
>>>>>>> master
=======
wrkqueues = {"orl_outbnd_cms":orl_outbnd_cms,"orl_ct":orl_ct,"orl_act":orl_act,"orl_cc":orl_cc,"spg_ct":spg_ct,"lv_outbnd_cms":lv_outbnd_cms,"lv_ct":lv_ct}
>>>>>>> parent of 55d8f87... a
=======
wrkqueues = {}
>>>>>>> parent of b8787b2... formatt

#SalesForce
orl_outbnd_sf = ["LOC-ORL-MKT-SalesForce", "SF-Orlando-Manual", "SF-RestrictDialing"] #Manual
spg_outbnd_sf = ["LOC-SPG-MKT-SalesForce", "SF-Springfield-Manual", "SF-RestrictDialing"]
lvn_outbnd_sf = ["LOC-LAS-MKT-SalesForce", "SF-Vegas-Manual", "SF-RestrictDialing"]

<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
licenses = [
    "Interaction Optimizer Client Access", "Interaction Optimizer Real-time Adherence Tracking",
    "Interaction Optimizer Schedulable",
    ]
=======
licenses = ["Interaction Optimizer Client Access", "Interaction Optimizer Real-time Adherence Tracking", "Interaction Optimizer Schedulable"]
>>>>>>> master
=======
licenses = ["Interaction Optimizer Client Access", "Interaction Optimizer Real-time Adherence Tracking", "Interaction Optimizer Schedulable"]
>>>>>>> parent of 55d8f87... a
=======
licenses = ["Interaction Optimizer Client Access", "Interaction Optimizer Real-time Adherence Tracking",
            "Interaction Optimizer Schedulable"]
>>>>>>> parent of b8787b2... formatt
roles = ["MKT-Agent", "MKT-SF-Agent", "MKT-CC-Agent"]

wb = load_workbook('excel_orgchart/orgchart.xlsx', read_only=True)
sh = wb['HGV_OrgChart']
ws = wb.active

column_header = {}

<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
=======
app = Application(backend='uia')
if serv == 1:
    p = pywinauto.findwindows.find_element(title="Interaction Administrator - [HiltonACD]")
else:
    p = pywinauto.findwindows.find_element(title="Interaction Administrator - [HiltonTCPA]")
app.connect(handle=p.handle)
if serv == 1:
    dlg = app.window(title="Interaction Administrator - [HiltonACD]")
else:
    dlg = app.window(title="Interaction Administrator - [HiltonTCPA]")
typein = app.dlg.type_keys
>>>>>>> parent of 55d8f87... a
#app.dlg.print_control_identifiers() #Check Identifiers


serv = input("1:CMS\n2:Salesforce\nSelect Option:")
serv = int(serv)

=======
=======
serv = input("1:CMS\n2:Salesforce\nSelect Option 1 or 2: ")
serv = int(serv)

>>>>>>> parent of b8787b2... formatt
app = Application(backend='uia')

if serv == 1:
    p = pywinauto.findwindows.find_element(title="Interaction Administrator - [HiltonACD]")
else:
    p = pywinauto.findwindows.find_element(title="Interaction Administrator - [HiltonTCPA]")
app.connect(handle=p.handle)
if serv == 1:
    dlg = app.window(title="Interaction Administrator - [HiltonACD]")
else:
    dlg = app.window(title="Interaction Administrator - [HiltonTCPA]")
typein = app.dlg.type_keys
#app.dlg.print_control_identifiers() #Check Identifiers

<<<<<<< HEAD

serv = input("1:CMS\n2:Salesforce\nSelect Option:")
serv = int(serv)

>>>>>>> master
location = input("Location orl, spg, lvn:")
=======
location = input("Select Location (orl, spg, lvn): ")
>>>>>>> parent of b8787b2... formatt
location = location.lower()

if serv == 1:
    if location == "spg":
        department = input("Select a department - ct:  ")
    if location == "lvn":
        department = input("Select a department - outbnd, ct: ")
    else:
        department = input("Select a department - outbnd, ct, act, cc: ")
    department = department.lower()

def main():
    get_alphabet()
    column_headers()

def Queues():
    orl_outbnd_cms = ["MKT-Outbound-Callback", "MKT-Outbound-Main2", "Orl_OUT_SUP"]
    orl_ct = ["CT Priority 1", "CT Priority 2", "LOC-ORL-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
    orl_act = ["LOC-ORL-MKT-ACT", "MKT-Activations-CallBack", "MKT-ACT-Main", "MKT-CC-BookDates",
               "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]
    orl_cc = ["LOC-ORL-MKT-CC", "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService",
              "MKT-CC-CustomerService-Priority2"]
    spg_ct = ["LOC-SPG-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
    lvn_outbnd_cms = ["LAS_OUT_SUP", "MKT-Outbound-Callback", "MKT-Outbound-Main2"]
    lvn_ct = ["CT Priority 1", "CT Priority 2", "LOC-LV-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]

    for i in ('orl_outbnd_cms', 'orl_ct', 'orl_act', 'orl_cc', 'spg_ct', 'lvn_outbnd_cms', 'lvn_ct'):
        wrkqueues[i] = locals()[i]

def get_alphabet():
    for c in ascii_lowercase:
        x = sh[c.upper() + "1"].value
        column_header[c.upper()] = sh[c.upper() + "1"].value

def getHeader(a):
    for k, v in column_header.items():
        if v == a:
            a = k
            return str(k)

#Find Letter associated with user details and add in excel
def column_headers():
    Add = "Add/Delete/Change/Transfer/Rehire"
    Add2 = "Add/Delete/Change"
    windows = "Windows for Adds"
    email = "Email Address"
    Name = "AgentName"
    cic_id = "CIC_ID"

<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
    orgchart_data(add, agent_username, agent_name, agent_tsr)
=======
    orgchart_data(add, agent_username,agent_name, agent_tsr)
>>>>>>> master
=======
    orgchart_data(add, agent_username,agent_name, agent_tsr)
>>>>>>> parent of 55d8f87... a
=======
    add = getHeader(Add) or getHeader(Add2)
    agent_username = getHeader(windows)
    agent_name = getHeader(Name)
    agent_tsr = getHeader(cic_id)
>>>>>>> parent of b8787b2... formatt

    orgchart_data(add, agent_username, agent_name, agent_tsr)

#Sends OrgChart Data to IA
def orgchart_data(add, windows, agent_name, agent_tsr):
    n = 2
    #Loop until last Row
    while n < sh.max_row:
        #Find All Adds
        if sh[add + str(n)].value == "Add":
            username = sh[windows + str(n)].value
            agentName = sh[agent_name + str(n)].value
            tsr = sh[agent_tsr + str(n)].value
<<<<<<< HEAD

            print("Adding User - " + username, agentName, tsr)
        
            try:
                Config(tsr, username)
                GetUserDetails(agentName)
                AutoACD()
                if serv == 1:
                    Roles(department)
                else:
                    SFRoles()
                if serv == 1:
                    getWorkGroups(location, department)
                    if department == "ct" or department == "cc":
                        Licensing()
                else:
                    getSFWorkGroups()
            except:
                print("User Already Exists " + username, agentName, tsr)
                if app.dlg["A User with that name already exists"].exists() == True:
                    app.dlg.OK.click_input()
                    app.dlg.Cancel.click_input()
            app.dlg.Cancel.click_input() #Change When Done
        n += 1

def getWorkGroups(loc, dept):
    if location == loc:
        if department == dept:
            if department == "outbnd":
                AgentWorkGroups(wrkqueues[loc + "_" + dept + "_" + "cms"])
            else:
                AgentWorkGroups(wrkqueues[loc + "_" + dept])
<<<<<<< HEAD

=======

<<<<<<< HEAD
            print("Adding User - " + username, agentName, tsr)
        
            try:
                Config(tsr, username)
                GetUserDetails(agentName)
                AutoACD()
                if serv == 1:
                    Roles(department)
                else:
                    SFRoles()
                if serv == 1:
                    getWorkGroups(location, department)
                    if department == "ct" or department == "cc":
                        Licensing()
                else:
                    getSFWorkGroups()
            except:
                print("User Already Exists " + username, agentName, tsr)
                if app.dlg["A User with that name already exists"].exists() == True:
                    app.dlg.OK.click_input()
                    app.dlg.Cancel.click_input()
            app.dlg.Cancel.click_input() #Change When Done
        n += 1

def getWorkGroups(loc, dept):
    if location == loc:
        if department == dept:
            if department == "outbnd":
                AgentWorkGroups(wrkqueues[loc + "_" + dept + "_" + "cms"])
            else:
                AgentWorkGroups(wrkqueues[loc + "_" + dept])
    
>>>>>>> master
=======
    
>>>>>>> parent of 55d8f87... a
def getSFWorkGroups(): #helper function here
    if location == "orl":
        AgentSFWorkGroups(orl_outbnd_sf)
    if location == "spg":
        AgentSFWorkGroups(spg_outbnd_sf)
    if location == "lvn":
        AgentSFWorkGroups(lv_outbnd_sf)

def Config(tsr, win_username):
    app.dlg.type_keys('^n')
    app.dlg.Edit0.type_keys(str(tsr) + "{ENTER}")
=======
            print(username, agentName, tsr)

            Config(tsr, username)
            GetUserDetails(agentName)
            AutoACD()

            if serv == 1:
                Roles(department)
            else:
                SFRoles()
            if serv == 1:
                getWorkGroups(location, department)
                if department == "ct" or department == "cc":
                    Licensing()
            else:
                getSFWorkGroups()
            app.dlg.Cancel.click_input() #Change When Done
        n += 1

def Config(tsr, win_username): #Check If Already Exist
    static = app.DialogName.child_window(title_re='.*Please contact your system administrator.',
                                         class_name_re='Static')
    """
    if static.exists(timeout=20): # if it opens no later than 20 sec.
        app.DialogName.OK.click()
    """
    app.dlg.type_keys('^n')
    app.dlg.Edit0.type_keys(str(tsr) + "{ENTER}")


>>>>>>> parent of b8787b2... formatt
    app.dlg.Edit3.type_keys("10102015") #Password
    app.dlg.Edit4.type_keys("10102015") #Confirm Password
    app.dlg.Edit7.type_keys("hgvcnt\\" + win_username) #Domain User
    getLoc()

def getLoc():
    app.dlg["ComboBox4"].click_input()
    if location == "orl":
        app.dlg["Orlando - Metro Center"].select()
    if location == "spg":
        app.dlg["Spingfield"].select()
    if location == "lvn":
        app.dlg["Las Vegas"].select()

#Personal Info
def GetUserDetails(agentName):
    name = agentName.split(" ")

    app.dlg.TabItem3.click_input()
    app.dlg.Edit0.type_keys(name[0]) #FirstName
    app.dlg.Edit2.type_keys(name[1]) #LastName
    app.dlg.Edit4.type_keys(name[0] + "{SPACE}" + name[1]) #Display Name

<<<<<<< HEAD
=======
def AutoACD():
    app.dlg.ACD.click_input()
    app.dlg.ListItem3.click_input()
    app.dlg.CheckBox0.click_input()

def Roles(deptmnt):
    app.dlg.Roles.click_input()
    app.dlg.Add.click_input()

    if deptmnt == "ct" or deptmnt == "outbnd":
        app.dlg.ListItem4.select()
    elif deptmnt == "act" or deptmnt == "cc":
        app.dlg.ListItem6.select()

    app.dlg.OK.click_input()

def SFRoles():
    app.dlg.Roles.click_input()
    app.dlg.Add.click_input()
    app.dlg.ListItem6.select()
    app.dlg.OK.click_input()
    
def getWorkGroups(loc, dept):
    if location == loc:
        if department == dept:
            if department == "outbnd":
                AgentWorkGroups(wrkqueues[loc + "_" + dept + "_" + "cms"])
            else:
                AgentWorkGroups(wrkqueues[loc + "_" + dept])

def getSFWorkGroups(): #helper function here
    if location == "orl":
        AgentSFWorkGroups(orl_outbnd_sf)
    if location == "spg":
        AgentSFWorkGroups(spg_outbnd_sf)
    if location == "lvn":
        AgentSFWorkGroups(lv_outbnd_sf)

>>>>>>> parent of b8787b2... formatt
#Assign CMS WorkGroups
def AgentWorkGroups(wrkgrps):
    num = 0
    app.dlg.Workgroups.click_input()
    app.dlg.OK.click_input()

    for x in wrkgrps:
        if location == "orl":
            if department == "outbnd":
                if num == 0:
                    ListBoxPos(-13)
                if num == 2:
                    ListBoxPos(-5)

            num += 1
            app.dlg[x].click_input()
            app.dlg.Add.click_input()

            if department == "ct":
                if num == 3:
                    ListBoxPos(-10)
            if department == "act" or department == "cc":
                if num == 1:
                    ListBoxPos(-6)

        if location == "spg":
            num += 1
            if department == "ct":
                if num == 1:
                    ListBoxPos(-2)
                if num == 2:
                    ListBoxPos(-8)

            app.dlg[x].click_input()
            app.dlg.Add.click_input()

        if location == "lvn":
            num += 1
            app.dlg[x].click_input()
            app.dlg.Add.click_input()

            if department == "outbnd":
                if num == 1:
                    ListBoxPos(-13)

            if department == "ct":
                if num == 3:
                    ListBoxPos(-10)

#Assign SF WorkGroups
def AgentSFWorkGroups(wrkgrps):
    num = 0
    app.dlg.Workgroups.click_input()
    app.dlg.OK.click_input()

    for x in wrkgrps:
        if location == "orl":
            num += 1
            app.dlg[x].click_input()
            app.dlg.Add.click_input()

        if location == "spg":
            num += 1
            app.dlg[x].click_input()
            app.dlg.Add.click_input()
            if num == 2:
                ListBoxPos(-2)

        if location == "lvn":
            num += 1
            app.dlg[x].click_input()
            app.dlg.Add.click_input()
            if num == 2:
                ListBoxPos(-2)


def ListBoxPos(scrollpos):
    app.dlg['ListBox'].click_input()
    app.dlg['ListBox'].wheel_mouse_input(wheel_dist=scrollpos)

def Licensing():
    app.dlg.Licensing.click_input()
    app.dlg.OK.click_input()
    app.dlg['Enable Licenses'].click_input()
    for ls in licenses:
        app.dlg[ls].type_keys("{SPACE}")
<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD

if serv == 1:
    get_IA_cms()
else:
    get_IA_manual()

=======
   
>>>>>>> master
=======
   
>>>>>>> parent of 55d8f87... a
=======

>>>>>>> parent of b8787b2... formatt
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
Workgroups - MKT-Outbound-Callback, MKT-Outbound-Main2
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
