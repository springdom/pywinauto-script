import time
import pywinauto
from pywinauto import application
from pywinauto.keyboard import SendKeys
from pywinauto import Desktop
from pywinauto.application import Application
from pywinauto.findwindows import find_window
from subprocess import Popen

#Display Modes
#app.dlg.print_control_identifiers() #Check Identifiers
opt1 = [1,2,1,2]
opt2 = [1,2,1,15,4,8]
opt3 = [1,2,1,15,4,10]

def main():
    print("Select From Menu")
    print("0: Check Availability")
    print("9: Check Tours")
    print("1: Lead Package Mktg Package Entry & Edit")
    print("2: Unlock Lead")
    print("3: Change to Kick")
    print("Already On Page?")
    print("4: Book Dates")
    print("5: Unlock Lead")
    print("6: Kick Package")
    print("")
    Sel = int(input("Select Option: "))
    print("")

    if Sel == 1:
        mode(opt1)
    elif Sel == 2:
        mode(opt2)
    elif Sel == 3:
        lead = input("Enter Lead: ")
        loc = input("Enter Location: ")
        change_to_kick(lead,loc)
    elif Sel == 4:
        lead = input("Enter Lead: ")
        loc = input("Enter Location: ")
        lead_package(lead,loc)
    elif Sel == 5:
        unlock_lead(lead,loc)
    elif Sel == 6:
        lead = input("Enter Lead: ")
        loc = input("Enter Location: ")
        kick_package(lead,loc)
    elif Sel == 0:
        check_dates()
    elif Sel == 9:
        check_tours()


#Window Settings
def set_window():
    app = Application(backend='uia')
    p = pywinauto.findwindows.find_element(best_match="AbsoluteTelnet")
    app.connect(handle=p.handle)
    dlg = app.window(best_match="AbsoluteTelnet")
    app.dlg.set_focus()


#Enter Mode
def mode(opt):
    set_window()
    SendKeys('{BACKSPACE 10}')
    SendKeys('{ENTER}')
    time.sleep(1)
    
    for num in opt:
        SendKeys(str(num))
        SendKeys('{ENTER}')
    SendKeys('{ENTER}')

def unlock_lead(leadn,loca):
    set_window()
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
    SendKeys('{TAB}' + office + '{ENTER}' + '{ENTER}' + '{ENTER 2}')
    toursa = input("Tours Available y or n: ")

    set_window()
    if toursa == "y":
        SendKeys('1' + '{ENTER 2}')
        SendKeys(manifest + '{ENTER 8}')
        

def check_avail():
    region = input("Enter Region: ")
    arrival = input("Enter Arrival: ")
    nights = input("Enter Nights: ")
    
    mode(opt1)
    SendKeys('{TAB}' + '2' + '{ENTER}')
    SendKeys(region + '{ENTER 2}')
    SendKeys(arrival + '{ENTER 2}')
    SendKeys(nights)
    SendKeys('{TAB}' + '3' + '{ENTER}')
    main()

def tour_booking():
    office = 0

def kick_package(leadn,loca):
    mode(opt1)
    
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
    SendKeys('{TAB}' + '{ENTER 8}')
    SendKeys('13')
    SendKeys('A')
    SendKeys('10' + '{ENTER 5}')
    SendKeys('Package Kicked as requested')
    
    
def change_to_kick(leadn,loca):
    mode(opt3)
   
    SendKeys(leadn + '{ENTER}')
    SendKeys(loca + '{ENTER}')
    SendKeys('{ENTER 2}')

    package_sel = input("Enter Package Number: ")
    set_window()
    
    SendKeys(package_sel + '{ENTER}')
    SendKeys("2")
    SendKeys("{ENTER}")
    SendKeys("9")
    SendKeys("{ENTER 2}")
    SendKeys("f")
    

def check_dates():
    region = input("Enter Region: ")
    arrival = input("Enter Arrival: ")
    
    mode(opt1)
    SendKeys('{TAB}' + '2' + '{ENTER}')
    SendKeys(region + '{ENTER 2}')
    SendKeys(arrival + '{ENTER}')
    SendKeys('{TAB}' + '2' + '{ENTER}')
    nory = input("Return To Options: y or n")
    main()

def check_tours():
    office = input("Enter Office: ")
    starting = input("Enter Tour Date: ")
    
    mode(opt1)
    SendKeys('{TAB}' + '3' + '{ENTER}')
    SendKeys(office + '{ENTER}')
    SendKeys(starting + '{ENTER 3}')
    SendKeys('{TAB}' + '2' + '{ENTER}')
    main()

main()
