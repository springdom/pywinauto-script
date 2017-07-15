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
print("")
Sel = int(input("Select Option: "))
print("")
opt1 = [1,2,1,2]
opt2 = [1,2,1,15,4,8]
opt3 = [1,2,1,15,4,10]



#app.dlg.print_control_identifiers() #Check Identifiers
#Window Settings
def set_window():
    app = Application(backend='uia')
    p = pywinauto.findwindows.find_element(best_match="AbsoluteTelnet")
    app.connect(handle=p.handle)
    dlg = app.window(best_match="AbsoluteTelnet")
    app.dlg.set_focus()


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
    set_window()
    SendKeys(leadn + '{ENTER}')
    SendKeys(loca + '{ENTER}')
    SendKeys('{ENTER 2}')
    package_sel = input("Select Package Number: ")
    
    set_window()
    
    SendKeys('{ENTER}')
    SendKeys(package_sel + '{ENTER}')
    SendKeys('7' + '{ENTER}' + 'A' + '{ENTER}')

    propn = input("Enter Property Number: ")
    unitt = input("Enter Unit Type: ")
    adults = input("Adults: ")
    kids = input("Kids: ")

    arrival = input("Arrival: ")
    nights = input("Nights: ") 

    set_window()
    
    SendKeys(propn + '{ENTER 2}')
    SendKeys(unitt + '{ENTER 4}')
    SendKeys(adults + '{ENTER 2}')
    SendKeys(kids + '{ENTER 2}')
    SendKeys(arrival + '{ENTER}')
    SendKeys(nights + '{ENTER 2}')
    
    bckspace = input("Go back? y or n:") # Backspace

    set_window()

    if bckspace == "y":
        SendKeys('{BACKSPACE}')

    proceed = input("Proceed - y or n: ")

    set_window()
    
    if proceed == 'y':
        SendKeys('{ENTER}')
    #Enter
    time.sleep(1)  
    office = input("Enter Office: ")
    manifest = input("Manifest: ")
    set_window()
    SendKeys('{TAB}' + office + '{ENTER}' + '{ENTER}' + '{ENTER}')
    


def tour_booking():
    office = 0

def kick_package(leadn,loca):
    set_window()
    
    SendKeys(leadn + '{ENTER}')
    SendKeys(loca + '{ENTER}')
    SendKeys('{ENTER 2}')
    package_sel = input("Select Package Number: ")
    
    set_window()
    
    SendKeys('{ENTER}')
    SendKeys(package_sel + '{ENTER}')
    SendKeys('20' + '{ENTER}')
    SendKeys('2' + '{ENTER}')
    SendKeys('14')
    SendKeys('{ENTER}')
    
    
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
    

if Sel == 1:
    mode(opt1)
    lead = input("Enter Lead: ")
    loc = input("Enter Location: ")
elif Sel == 2:
    mode(opt2)
    lead = input("Enter Lead: ")
    loc = input("Enter Location: ")
elif Sel == 3:
    mode(opt3)
elif Sel == 4:
    lead = input("Enter Lead: ")
    loc = input("Enter Location: ")
    lead_package(lead,loc)
elif Sel == 5:
    unlock_lead(lead,loc)



