"""
Automates Interaction Administrator
"""
from string import ascii_lowercase
import pywinauto
from pywinauto import application
#from pywinauto.keyboard import SendKeys
#from pywinauto import Desktop
from pywinauto.application import Application
#from pywinauto.findwindows import find_window
from openpyxl import load_workbook

<<<<<<< HEAD
=======
serv = input("1:CMS\n2:Salesforce\nSelect Option: ")
serv = int(serv)

>>>>>>> Update CreateUser.py
#CMS
orl_outbnd_cms = ["MKT-Outbound-Callback", "MKT-Outbound-Main2", "Orl_OUT_SUP"]
orl_ct = ["CT Priority 1", "CT Priority 2", "LOC-ORL-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
orl_act = [
    "LOC-ORL-MKT-ACT", "MKT-Activations-CallBack", "MKT-ACT-Main", "MKT-CC-BookDates",
    "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2",
    ]

orl_cc = [
    "LOC-ORL-MKT-CC", "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1",
    "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2",
          ]

spg_ct = ["LOC-SPG-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
lvn_outbnd_cms = ["LAS_OUT_SUP", "MKT-Outbound-Callback", "MKT-Outbound-Main2"]
lv_ct = ["CT Priority 1", "CT Priority 2", "LOC-LV-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
<<<<<<< HEAD
<<<<<<< HEAD
wrkqueues = {"orl_outbnd_cms":orl_outbnd_cms,"orl_ct":orl_ct,"orl_act":orl_act,"orl_cc":orl_cc,"spg_ct":spg_ct,"lvn_outbnd_cms":lvn_outbnd_cms,"lvn_ct":lv_ct}
=======
wrkqueues = {"orl_outbnd_cms":orl_outbnd_cms,"orl_ct":orl_ct,"orl_act":orl_act,"orl_cc":orl_cc,"spg_ct":spg_ct,"lv_outbnd_cms":lv_outbnd_cms,"lv_ct":lv_ct}
>>>>>>> Update CreateUser.py
=======
wrkqueues = {
    "orl_outbnd_cms":orl_outbnd_cms,"orl_ct":orl_ct,"orl_act":orl_act,
    "orl_cc":orl_cc, "spg_ct":spg_ct, "lvn_outbnd_cms":lvn_outbnd_cms, "lvn_ct":lv_ct,
    }
>>>>>>> 8e45f202e44ca3f92e7df0b9aeb2e201bf981abf

#SalesForce
orl_outbnd_sf = ["LOC-ORL-MKT-SalesForce", "SF-Orlando-Manual", "SF-RestrictDialing"] 
spg_outbnd_sf = ["LOC-SPG-MKT-SalesForce", "SF-Springfield-Manual", "SF-RestrictDialing"]
lvn_outbnd_sf = ["LOC-LAS-MKT-SalesForce", "SF-Vegas-Manual", "SF-RestrictDialing"]

licenses = [
    "Interaction Optimizer Client Access", "Interaction Optimizer Real-time Adherence Tracking",
    "Interaction Optimizer Schedulable",
    ]

roles = ["MKT-Agent", "MKT-SF-Agent", "MKT-CC-Agent"]

<<<<<<< HEAD
<<<<<<< HEAD
wb = load_workbook('excel_orgchart/orgchart.xlsx', read_only=True)
=======
wb = load_workbook('orgchart.xlsx', read_only=True)
>>>>>>> parent of c3fbc65... added build
sh = wb['HGV_OrgChart']
ws = wb.active

<<<<<<< HEAD
serv = input("1:CMS\n2:Salesforce\nSelect Option:")
=======
wb = load_workbook('excel_orgchart\orgchart.xlsx', read_only=True)
sh = wb['HGV_OrgChart']
ws = wb.active

<<<<<<< HEAD
serv = input("1:CMS\n2:Salesforce\nSelect Option 1 or 2: ")
<<<<<<< HEAD
>>>>>>> 8e45f202e44ca3f92e7df0b9aeb2e201bf981abf
=======
=======
serv = input("1:CMS\n2:Salesforce\nSelect Option:")
>>>>>>> b8787b204b817353efab1140cdc4f518a4c0ea11
>>>>>>> cleanup
serv = int(serv)

=======
>>>>>>> Update CreateUser.py
column_header = {}
<<<<<<< HEAD

=======
>>>>>>> cleanup
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
<<<<<<< HEAD
<<<<<<< HEAD
location = input("Location orl, spg, lvn:")
location = location.lower()

=======
>>>>>>> 8e45f202e44ca3f92e7df0b9aeb2e201bf981abf
=======
=======
location = input("Location orl, spg, lvn:")
location = location.lower()

>>>>>>> b8787b204b817353efab1140cdc4f518a4c0ea11
>>>>>>> cleanup
if serv == 1:
    if location == "spg":
        department = input("Select a department - ct:  ")
    elif location == "lvn":
<<<<<<< HEAD
    if location == "lvn":
=======
>>>>>>> b8787b204b817353efab1140cdc4f518a4c0ea11
        department = input("Select a department - outbnd, ct: ")
    else:
        department = input("Select a department - outbnd, ct, act, cc: ")
    department = department.lower()

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

#app.dlg.print_control_identifiers() #Check Identifiers

location = input("Location orl, spg, lvn: ")
location = location.lower()
if serv == 1:
    department = input("Department outbnd, ct, act, cc: ")
    department = department.lower()
>>>>>>> Update CreateUser.py

=======
>>>>>>> 8e45f202e44ca3f92e7df0b9aeb2e201bf981abf
def main():
    get_alphabet()
    column_headers()

def get_alphabet():
    for alpha in ascii_lowercase:
        x = sh[alpha.upper() + "1"].value
        column_header[alpha.upper()] = sh[alpha.upper() + "1"].value

def sed(a):
    for k, v in column_header.items():
        if v == a:
            a = k
            return str(k)

def column_headers():
    Add = "Add/Delete/Change/Transfer/Rehire"
    Add2 = "Add/Delete/Change"
    windows = "Windows for Adds"
    email = "Email Address"
    Name = "AgentName"
    cic_id = "CIC_ID"
    
    add = sed(Add) or sed(Add2)
    agent_username = sed(windows)
    agent_name = sed(Name)
    agent_tsr = sed(cic_id)
<<<<<<< HEAD

    orgchart_data(add, agent_username,agent_name, agent_tsr)
=======
>>>>>>> b8787b204b817353efab1140cdc4f518a4c0ea11

<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD

=======
>>>>>>> Update CreateUser.py
=======
    add = getHeader(Add) or getHeader(Add2)
    agent_username = getHeader(windows)
    agent_name = getHeader(Name)
    agent_tsr = getHeader(cic_id)

    orgchart_data(add, agent_username, agent_name, agent_tsr)

#Sends OrgChart Data to IA
>>>>>>> 8e45f202e44ca3f92e7df0b9aeb2e201bf981abf
=======
    orgchart_data(add, agent_username,agent_name, agent_tsr)

>>>>>>> cleanup
def orgchart_data(add, windows, agent_name, agent_tsr):
    n = 2
    while n < sh.max_row:
        if sh[add + str(n)].value == "Add":
            username = sh[windows + str(n)].value
            agentName = sh[agent_name + str(n)].value
            tsr = sh[agent_tsr + str(n)].value
<<<<<<< HEAD

<<<<<<< HEAD
=======
>>>>>>> 8e45f202e44ca3f92e7df0b9aeb2e201bf981abf
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
<<<<<<< HEAD
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
            app.dlg.Cancel.click_input() #Change When DOne
>>>>>>> Update CreateUser.py
        n += 1
=======
        n += 1 bv
>>>>>>> 8e45f202e44ca3f92e7df0b9aeb2e201bf981abf

def getWorkGroups(loc, dept):
    if location == loc:
        if department == dept:
            if department == "outbnd":
                AgentWorkGroups(wrkqueues[loc + "_" + dept + "_" + "cms"])
            else:
                AgentWorkGroups(wrkqueues[loc + "_" + dept])
    
<<<<<<< HEAD
def getSFWorkGroups(): #helper function here
=======
def getSFWorkGroups(): # helper function here
>>>>>>> Update CreateUser.py
    if location == "orl":
        AgentSFWorkGroups(orl_outbnd_sf)
    if location == "spg":
        AgentSFWorkGroups(spg_outbnd_sf)
    if location == "lvn":
        AgentSFWorkGroups(lv_outbnd_sf)

def Config(tsr, win_username, passwd = "1010215", domain = "hgvcnt\\"):

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
        n += 1 bv

def getWorkGroups(loc, dept):
    if location == loc:
        if department == dept:
            if department == "outbnd":
                AgentWorkGroups(wrkqueues[loc + "_" + dept + "_" + "cms"])
            else:
<<<<<<< HEAD
                getSFWorkGroups()
            app.dlg.Cancel.click_input() #Change When Done
        n += 1
=======
                AgentWorkGroups(wrkqueues[loc + "_" + dept])
    
def getSFWorkGroups(): #helper function here
    if location == "orl":
        AgentSFWorkGroups(orl_outbnd_sf)
    if location == "spg":
        AgentSFWorkGroups(spg_outbnd_sf)
    if location == "lvn":
        AgentSFWorkGroups(lv_outbnd_sf)

def Config(tsr, win_username, passwd = "1010215", domain = "hgvcnt\\"):
    app.dlg.type_keys('^n')
    app.dlg.Edit0.type_keys(str(tsr) + "{ENTER}")
    app.dlg.Edit3.type_keys(passwd) #Password
    app.dlg.Edit4.type_keys(passwd) #Confirm Password
    app.dlg.Edit7.type_keys(domain + win_username) #Domain User
    getLoc()
>>>>>>> b8787b204b817353efab1140cdc4f518a4c0ea11

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

<<<<<<< HEAD
=======
def getSFWorkGroups(): #helper function here
    if location == "orl":
        AgentSFWorkGroups(orl_outbnd_sf)
    if location == "spg":
        AgentSFWorkGroups(spg_outbnd_sf)
    if location == "lvn":
        AgentSFWorkGroups(lv_outbnd_sf)

>>>>>>> 8e45f202e44ca3f92e7df0b9aeb2e201bf981abf
#Assign CMS WorkGroups
=======

>>>>>>> cleanup
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

def Licensing():
    app.dlg.Licensing.click_input()
    app.dlg.OK.click_input()
    app.dlg['Enable Licenses'].click_input()
<<<<<<< HEAD

    for Ls in licenses:
        app.dlg[Ls].type_keys("{SPACE}")

=======
    for Ls in licenses:
        app.dlg[Ls].type_keys("{SPACE}")
   
>>>>>>> b8787b204b817353efab1140cdc4f518a4c0ea11
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
