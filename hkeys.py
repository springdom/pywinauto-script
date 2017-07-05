import time
import pywinauto
from pywinauto import application
from pywinauto.keyboard import SendKeys
from pywinauto import Desktop
from pywinauto.application import Application
from pywinauto.findwindows import find_window
from subprocess import Popen

#Display Modes
print("Select From Menu")
print("1: Lead Package Mktg Package Entry & Edit")
print("2: Unlock Lead")
print("Already On Page?")
print("3: Book Dates")
print("4: Unlock Lead")

Sel = int(input("Select Option: "))

opt1 = [1,2,1,2]
opt2 = [1,2,1,15,4,8]

lead = input("Enter Lead: ")
loc = input("Enter Location: ")

app = Application(backend='uia')
p = pywinauto.findwindows.find_element(best_match="AbsoluteTelnet")
app.connect(handle=p.handle)
dlg = app.window(best_match="AbsoluteTelnet")

app.dlg.set_focus()
#app.dlg.print_control_identifiers() #Check Identifiers


#Enter Mode
def mode(opt):
    for num in opt:
        SendKeys(str(num))
        SendKeys('{ENTER}')
    SendKeys('{ENTER 2}')

def unlock_lead(leadn,loca):
    SendKeys(leadn + '{ENTER}')
    SendKeys(loca + '{ENTER}')
    SendKeys('{ENTER 4}')

#Lead Mktg Package Entry & Edit
def lead_package(leadn,loca):
    SendKeys(leadn + '{ENTER}')
    SendKeys(loca + '{ENTER}')
    SendKeys('{ENTER 2}')
    #Ask Which Package
    #SK-7,SK-A
    #1stSK-PropName,

print(Sel)
if Sel == 1:
    mode(opt1)
elif Sel == 2:
    mode(opt2)
elif Sel == 3:
    lead_package(lead,loc)
elif Sel == 4:
    unlock_lead(lead,loc)

Book Dates
    #Recent Date Opt
    #Opt 7, Opt A
    #Property
    #Enter 2
    #Unit Type
    #Enter 4
    #Adults,Enter, Children, ENter
    #Arrival,Enter,Nights,Enter
    #Ask if Available, Enter
    #Ask if Backspace, if yes Backspace
    #enter, ThreadSleep
    ###Tour
    #Tab,office,Enter num,Enter 3


