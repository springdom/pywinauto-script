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
print("3: Change to Kick")
print("Already On Page?")
print("4: Book Dates")
print("5: Unlock Lead")

Sel = int(input("Select Option: "))

opt1 = [1,2,1,2]
opt2 = [1,2,1,15,4,8]
opt3 = [1,2,1,15,4,10]


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
"""
    package_sel = int(input("Select Package Number: "))
    prop = int(input("Enter Property Number: "))
    unit_type int(input("Unit Type: "))
    adults = int(input("How many adults: "))
    children = int(input("How many children"))
    arrival = input("Arrival: ")
    nights = input("Nights: ")
    availability = 'n'
    backspace = 'n'
""" 
    #Ask Which Package
    #SK-7,SK-A
    #1stSK-PropName,

def tour_booking():
    office = 0
    
def change_to_kick(leadn,loca):
    SendKeys(leadn + '{ENTER}')
    SendKeys(loca + '{ENTER}')
    SendKeys('{ENTER 2}')

    package_sel = input("Enter Package Number: ")
    SendKeys(package_sel)
    SendKeys("2")
    SendKeys("{ENTER}")
    SendKeys("9")
    SendKeys("{ENTER 2}")
    SendKeys("f")
    

print(Sel)
if Sel == 1:
    mode(opt1)
elif Sel == 2:
    mode(opt2)
elif Sel == 3:
    mode(opt3)
elif Sel == 4:
    lead_package(lead,loc)
elif Sel == 5:
    unlock_lead(lead,loc)

###Book Dates
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


