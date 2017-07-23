import time
import win32clipboard
import pywinauto
from pywinauto import application
from pywinauto.keyboard import SendKeys
from pywinauto import Desktop
from pywinauto.application import Application
from pywinauto.findwindows import find_window

#app.dlg.print_control_identifiers() #Check Identifiers

opt1 = [1,2,1,2]
opt2 = [1,2,1,15,4,8]
opt3 = [1,2,1,15,4,10]

def main():
    print("-" * 20)
    print("Select From Menu")
    print("0: Check Availability")
    print("9: Check Tours")
    print("8: Check Avail")
    print("7: Cashback Check")
    print("-" * 20)
    print("1: Book Dates")
    print("2: Unlock Lead")
    print("-" * 20)
    print("3: Kick Package")
    print("4: Change to Kick")
    print("-" * 20)
    Sel = int(input("Select Option: "))
    print("")

    if Sel >= 1 and Sel <= 3 or Sel == 7:
        lead = input("Enter Lead: ")
        loc = input("Enter Location: ")

    if Sel == 1:
        check_avail()
        time.sleep(2)
        lead_package(lead,loc)
    elif Sel == 2:
        unlock_lead(lead,loc)
    elif Sel == 3:
        kick_package(lead,loc)
    elif Sel == 4:
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        leadvar = data.split("-")
        change_to_kick(leadvar[1],leadvar[0])
    elif Sel == 0:
        check_dates()
    elif Sel == 9:
        check_tours()
    elif Sel == 8:
        check_avail()
        main()
    elif Sel == 7:
        check_cashback(lead,loc)
    else:
        print("Invalid Option")
        main()

#Window Settings
def set_window():
    
    app = Application(backend='uia')
    p = pywinauto.findwindows.find_element(best_match="AbsoluteTelnet")
    app.connect(handle=p.handle)
    dlg = app.window(best_match="AbsoluteTelnet")
    app.dlg.set_focus()
    #app.dlg.type_keys("Test")

#Enter Mode
def mode(opt):
    set_window()
    SendKeys('{BACKSPACE 10}')
    SendKeys('{ENTER}')
    time.sleep(1.5)
    
    for num in opt:
        SendKeys(str(num))
        SendKeys('{ENTER}')
    SendKeys('{ENTER}')

#Lead Mktg Package Entry & Edit
def lead_package(leadn,loca):
    mode(opt1)
    
    SendKeys(leadn + '{ENTER}')
    SendKeys(loca + '{ENTER}')
    SendKeys('{ENTER 2}')
    #check if lead good
    package_sel = input("Select Package Number: ")
    
    set_window()
    
    SendKeys('{ENTER}')
    SendKeys(package_sel + '{ENTER}')
    #check if package in use
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
    
    bckspace = input("Go back? y or n: ") # Backspace

    set_window()

    if bckspace == "y":
        SendKeys('{BACKSPACE}')

    proceed = input("Proceed - y or n: ")

    set_window()
    
    if proceed == 'y':
        SendKeys('{ENTER}')

    time.sleep(1)  
    office = input("Enter Office: ")
    manifest = input("Manifest: ")
    set_window()
    SendKeys('{TAB}' + office + '{ENTER}' + '{ENTER}' + '{ENTER 2}')
    toursa = input("Tours Available y or n: ")

    set_window()
    
    if toursa == "y":
        SendKeys('1' + '{ENTER 2}')
        SendKeys(manifest)
        SendKeys('{ENTER 10}')
        time.sleep(1.30)
        SendKeys('{ENTER 15}')
        SendKeys('f' + '{ENTER}')
        SendKeys('y' + "{ENTER}")
        amount = input("Enter amount: ")
        set_window()
        SendKeys(amount)
    else:
        pass
    main()

#Unlock Lead
def unlock_lead(leadn,loca):
    mode(opt2)
    set_window()
    SendKeys(leadn + '{ENTER}')
    SendKeys(loca + '{ENTER}')
    SendKeys('{ENTER 4}')
    main()

#Kick Package
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
    SendKeys('14' + '{ENTER}')
    SendKeys('{TAB}')
    time.sleep(3)
    SendKeys('{ENTER 8}')
    SendKeys('13')
    #SendKeys('A')
    #SendKeys('10' + '{ENTER 5}')
    #SendKeys('Package Kicked as requested')
    main()
    
#Change To Kick
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
    SendKeys("f" + "{ENTER}")
    main()

#Check Availability
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


#Check Dates
def check_dates():
    region = input("Enter Region: ")
    arrival = input("Enter Arrival: ")
    
    mode(opt1)
    SendKeys('{TAB}' + '2' + '{ENTER}')
    SendKeys(region + '{ENTER 2}')
    SendKeys(arrival + '{ENTER}')
    SendKeys('{TAB}' + '2' + '{ENTER}')
    main()

#Check Tours
def check_tours():
    office = input("Enter Office: ")
    starting = input("Enter Tour Date: ")
    
    mode(opt1)
    SendKeys('{TAB}' + '3' + '{ENTER}')
    SendKeys(office + '{ENTER}')
    SendKeys(starting + '{ENTER 3}')
    SendKeys('{TAB}' + '2' + '{ENTER}')
    main()

def check_cashback(leadn,loca):
    mode(opt3)

    SendKeys(leadn + '{ENTER}')
    SendKeys(loca + '{ENTER}')
    SendKeys('{ENTER 2}')
    package_sel = input("Enter Package Number: ")
    set_window()
    SendKeys(package_sel + '{ENTER}')
    SendKeys('18' + '{ENTER}')
    time.sleep(1)
    SendKeys('{BACKSPACE 3}')
    main()    

#Build A Package
def build_package():
    mode(opt1)
    #Check if Exist first

    lastName = input("Enter Last Name: ")
    FirstName = input("Enter First Name: ")
    PhoneNo = input("Phone Number: ")
    phoneNo2 = input("Second Phone Number: ")
    Country = input("Enter Country Code: ")
    addr = input("Address: ")
    zipCode = input("Input ZipCode: ")
    hhn = input("Hilton Honors Number: ")
    email = input("Enter Email: ")
    lang = input("Preferred Language: ")
    natio = input("Nationality: ")
    
    #LastName,FirstName,Phone Number
    SendKeys(LastName + "{Enter}")
    SendKeys(FirstName + "{Enter}")
    #Country,Enter*4
    SendKeys(Country + "{ENTER 4}")
    #zip code,address,Enter 5
    SendKeys(zipCode + "{ENTER}")
    SendKeys(addr + "{ENTER 5}")
    #Enter 9
    SendKeys("{ENTER 9}")
    #HHn Enter 2,email enter 2,
    SendKeys(hhn + "{ENTER 2}")
    SendKeys(email + "{ENTER 2}")
    #pref lang, Nationailty enter 2
    SendKeys(lang + "{ENTER}")
    SendKeys(natio + "{ENTER}")
    
def is_good(x):
    pass

main()
