import time
import pywinauto
from pywinauto.keyboard import SendKeys
from pywinauto.application import Application
from openpyxl import load_workbook
#from pywinauto import application
#import win32clipboard
#from pywinauto import Desktop
#from pywinauto.findwindows import find_window

#app.dlg.print_control_identifiers() #Check Identifiers
wb = load_workbook('KeyCodes.xlsx', data_only=True)
sh = wb["Sheet1"]
ws = wb.active

wb2 = load_workbook('changetokick.xlsx', data_only=True)
sh2 = wb["Sheet1"]
ws2 = wb2.active

outbnd_orl = {}
orl_call_trsf = {}
lvn_call_trsf = {}
gldmtn_call_trsf = {}

opt1 = [1, 2, 1, 2]
opt2 = [1, 2, 1, 15, 4, 8]
opt3 = [1, 2, 1, 15, 4, 10]

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
    print("5: Create Package")
    print("-" * 20)
    Sel = int(input("Select Option: "))
    print("")

    if Sel >= 1 and Sel <= 3 or Sel == 7:
        try:
            lead = input("Enter Lead: ")
            loc = input("Enter Location: ")
        except (KeyboardInterrupt, SystemExit):
            main()

    if Sel == 1:
        check_avail()
        checkavail = input("Available: y or n: ")
        if checkavail == "y":
            lead_package(lead, loc)
        else:
            main()
    elif Sel == 2:
        unlock_lead(lead, loc)
    elif Sel == 3:
        kick_package(lead, loc)
    elif Sel == 4:
        """
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        leadvar = data.split("-")
        change_to_kick(leadvar[1],leadvar[0])
        """
        kick_list(3, 1, 56)
    elif Sel == 5:
        build_package()
    elif Sel == 0:
        check_dates()
    elif Sel == 9:
        check_tours()
    elif Sel == 8:
        check_avail()
        main()
    elif Sel == 7:
        check_cashback(lead, loc)
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

#KeyCodes - Outbound, Orlando, Las Vegas, Gold Mountain
def keycodes(m_row, mx_col, mx_row, x):
    num = 0
    for row in ws.iter_rows(min_row=m_row, max_col=mx_col, max_row=mx_row):
        for cell in row:
            num += 1
            if num == 1:
                #print(cell.value)
                x[cell.value.lower()] = None
            if m_row > 17:
                if num == 33:
                    num = 0
            else:
                if num == 3:
                    num = 0

def assign_key(n, x):
    if n > 17:
        for k, v in x.items():
            x[k] = sh['D%s' % n].value
            n += 1
    else:
        for k, v in x.items():
            x[k] = sh['B%s' % n].value
            n += 1

def kick_list(m_row, mx_col, mx_row):
    for row in ws2.iter_rows(min_row=m_row, max_col=mx_col, max_row=mx_row):
        for cell in row:
            lead = cell.value.split("-")
            change_to_kick(lead[1], lead[0])

#Lead Mktg Package Entry & Edit
def lead_package(leadn, loca):
    mode(opt1)
    SendKeys(leadn + '{ENTER}')
    SendKeys(loca + '{ENTER}')
    SendKeys('{ENTER 2}')
    try:
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
            SendKeys('{ENTER 9}')
            time.sleep(1.30)
            SendKeys('{ENTER 17}')
            SendKeys('f' + '{ENTER}')
            SendKeys('y' + "{ENTER}")
            amount = input("Enter amount: ")
            set_window()
            SendKeys(amount)
            SendKeys("{ENTER}" + "n")
        else:
            pass
        main()
    except (KeyboardInterrupt, SystemExit):
        main()

#Unlock Lead
def unlock_lead(leadn, loca):
    mode(opt2)
    set_window()
    SendKeys(leadn + '{ENTER}')
    SendKeys(loca + '{ENTER}')
    SendKeys('{ENTER 5}')
    main()

#Kick Package
def kick_package(leadn, loca):
    mode(opt1)
    txtcmnt = "Package Kicked as requested"
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
    SendKeys('{ENTER 10}')
    SendKeys('13')
    #SendKeys('A')
    #SendKeys('10' + '{ENTER 5}')
    #SendKeys(txtcmnt.replace(" ","{SPACE}") + "{ENTER}")
    main()

#Change To Kick
def change_to_kick(leadn, loca):
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
    #main()

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

def check_cashback(leadn, loca):
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
    SendKeys("{ENTER}")
    last_name = input("Enter Last Name: ")
    first_name = input("Enter First Name: ")
    phone_no = input("Phone Number: ")
    phone_no2 = input("Second Phone Number: ")
    country = input("Enter Country Code: ")
    addr = input("Address: ")
    zip_code = input("Input Zip Code: ")
    hhn = input("Hilton Honors Number: ")
    email = input("Enter Email: ")
    lang = input("Preferred Language: ")
    natio = input("Nationality: ")

    set_window()

    #last_name,first_name,Phone Number
    SendKeys(last_name.replace(" ", "{SPACE}") + "{ENTER}")
    SendKeys(first_name.replace(" ", "{SPACE}") + "{ENTER}")
    SendKeys(phone_no + "{ENTER 2}")
    #country,Enter*4
    if "/" in last_name or "/" in first_name:
        SendKeys(country + "{ENTER 5}")
    else:
        SendKeys(country + "{ENTER 4}")
    SendKeys(zip_code + "{ENTER}")
    SendKeys(addr.replace(" ", "{SPACE}") + "{ENTER 5}")
    time.sleep(2)

    #Check if address good
    addrgood = input("Select Option: override continue ")

    set_window()
    if addrgood == "o":
        SendKeys("o" + "{ENTER}")
    elif addr == "c":
        pass

    SendKeys(phone_no2)
    SendKeys("{ENTER 9}")
    SendKeys(hhn + "{ENTER 2}")
    SendKeys(email + "{ENTER 2}")
    SendKeys(lang + "{ENTER}")
    SendKeys(natio + "{ENTER 2}")

    ###Marketing Key Section
    ##Check Department
    dptm = input("Department outbnd, orl, gldmtn, lvn: ")
    key_loca = input("Location: ")
    promo = input("POS promo: ")
    pckgcode = input("Package Code: ")
    stt = input("Marital Status: ")
    tsr = input("Agent tsr: ")
    if dptm == "outbnd":
        keycodes(4, 3, 12, outbnd_orl)
        assign_key(4, outbnd_orl)
        mktkey = outbnd_orl[key_loca]
    if dptm == "orl":
        keycodes(18, 33, 27, orl_call_trsf)
        assign_key(18, orl_call_trsf)
        mktkey = orl_call_trsf[key_loca]
    if dptm == "gldmtn":
        keycodes(46, 33, 55, gldmtn_call_trsf)
        assign_key(46, gldmtn_call_trsf)
        mktkey = gldmtn_call_trsf[key_loca]
    if dptm == "lvn":
        keycodes(32, 33, 41, lvn_call_trsf)
        assign_key(32, lvn_call_trsf)
        mktkey = lvn_call_trsf[key_loca]

    set_window()
    SendKeys(mktkey + "{ENTER}")
    SendKeys("{BACKSPACE}")
    SendKeys("{ENTER 2}")
    SendKeys(promo + "{ENTER 6}")

    #Payment
    SendKeys(pckgcode + "{ENTER 5}")
    price = input("Price: ")
    set_window()
    SendKeys(str(price) + "{ENTER 3}")
    ##add CC
    #is cc good then
    ccgood = input("CC accepted: ")

    if not ccgood:
        main()

    set_window()
    SendKeys("{ENTER 2}")
    SendKeys(str(tsr) + "{ENTER}")
    SendKeys(str(tsr) + "{ENTER}")
    SendKeys(str(tsr) + "{ENTER 12}")

    SendKeys("1" + "{ENTER}")
    SendKeys("1" + "{ENTER}")
    SendKeys(mktkey + "{ENTER 3}")
    SendKeys(stt + "{ENTER 10}")
    main()

main()