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

"""
Check Location
Check Dept
Check Type
"""

#app.dlg.print_control_identifiers() #Check Identifiers
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
app.dlg.Edit.type_keys("Matthew") #Extension
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
app.dlg.Edit0.type_keys("Dave")
app.dlg.Edit2.type_keys("Sherman")
app.dlg.Edit4.type_keys("Dave{SPACE}Sherman")

#ACD
"""
Auto ACD
"""
app.dlg.ACD.click_input()
app.dlg.ListItem3.click_input()
app.dlg.CheckBox0.click_input()

#Roles
"""
Based on Dept
"""
app.dlg.Roles.click_input()
app.dlg.Add.click_input()
app.dlg.ListItem4.select()
app.dlg.OK.click_input() #Multiple Select?

#WorkGroup
"""
Check Dept
"""
app.dlg.Workgroups.click_input()
app.dlg.OK.click_input()
app.dlg.ListItem4.click_input()
app.dlg['CT Priority 1'].click_input()
app.dlg['CT Priority 2'].click_input()
app.dlg.Add.click_input()

#Licensing

app.dlg.Licensing.click_input()
app.dlg.OK.click_input()
app.dlg['Enable Licenses'].click_input()
app.dlg['Interaction Optimizer Client Access'].click_input()
app.dlg['Interaction Optimizer Client Access'].type_keys("{SPACE}")
app.dlg['Interaction Optimizer Real-time Adherence Tracking'].click_input()
app.dlg['Interaction Optimizer Real-time Adherence Tracking'].type_keys("{SPACE}")
app.dlg['Interaction Optimizer Schedulable'].click_input()
app.dlg['Interaction Optimizer Schedulable'].type_keys("{SPACE}")
app.dlg.OK.click_input()
