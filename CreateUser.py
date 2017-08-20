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


"""
TSR
Extension
PassWord
Confirm Password
NT Domain User
"""

#Configuration
app.dlg.menu_select("File->New")
app.dlg.Edit0.type_keys("123456" + "{ENTER}")
app.dlg.Edit.type_keys("Matthew")
app.dlg.Edit3.type_keys("10102015")
app.dlg.Edit4.type_keys("10102015")
app.dlg.Edit7.type_keys("hgvcnt\\")


#Personal Info
app.dlg.TabItem3.click_input()

"""
#WorkGroups
app.dlg.TabItem4.click_input()

#Licensing
app.dlg.TabItem2.click_input()

#Roles
app.dlg.TabItem5.click_input()

#ACD
app.dlg.TabItem7.click_input()
"""
