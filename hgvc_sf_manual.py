from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time

#Select Region - est/cnt/pcf
print("est - Eastern, cnt - Central, pcf - Pacific")
region = input("Enter Region Name: ")

#Select cell 1 or cell 2
print("Enter 1 for Cell 1 or 2 for Cell 2")
cellopt = int(input("Enter Whether Leads are for Cell 1 or Cell 2 agents: "))
#amount = input("Amount of leads to add(10, 25, 50, 100, 200): ")

driver = webdriver.Firefox()
by_name = driver.find_element_by_name
by_id = driver.find_element_by_id

user_name = "username"
passwd = "password"

selector = "00B38000007XnvD_paginator_rpp_target"

def BrowserSetup():
    driver.get("https://identitynow.hgv.com/")
    #driver.get("https://hgv.my.salesforce.com/home/home.jsp?tsid=02u38000000NcOF")
    #assert "HGV IdentityNow" in driver.title
    username = by_id('username')
    password = by_id('password')
    username.send_keys(user_name)
    password.send_keys(passwd)

def startUp():
    by_id("signIn").click()
    time.sleep(5)
    driver.get("https://hgv.my.salesforce.com/home/home.jsp?tsid=02u38000000NcOF")
    time.sleep(7)
    by_id("01r38000000Hv1x_Tab").click()
    time.sleep(2)
    SelectQueue(selector)
    
    try:
        by_name("go").click()
    except:
        pass
    
    time.sleep(3)

def leadOpt():
    selEastern =  "00B38000007Xnw2_paginator_rpp_target"
    selCentral =  "00B38000007Xnvr_paginator_rpp_target"
    selPacific =  "00B38000007XnwQ_paginator_rpp_target"

    #Cell2
    selEastern2 = "00B38000007Xnw1_paginator_rpp_target"
    selCentral2 = "00B38000007Xnvw_paginator_rpp_target"
    selPacific2 = "00B38000007XnwV_paginator_rpp_target"

    #Available & Callable
    selEasternAC = "00B38000007XnvD_paginator_rpp_target"
    selCentralAC = "00B38000007Xnv8_paginator_rpp_target"
    selPacificAC = "00B38000007Xnvh_paginator_rpp_target"

    #Cell2 Available & Callable
    selEasternAC2 = "00B38000007XnvD_paginator_rpp_target"
    selCentralAC2 = "00B38000007Xnv8_paginator_rpp_target"
    selPacificAC2 = "00B38000007Xnvh_paginator_rpp_target"
    
    if cellopt == 1:
        if region == "est":
            add_leads(selEasternAC)
        elif region == "cnt":
            add_leads(selCentralAC)
        elif region == "pcf":
            add_leads(selPacificAC)
    elif cellopt == 2:
        if region == "est":
            add_leads(selEasternAC2)
        elif region == "cnt":
            add_leads(selCentralAC2)
        elif region == "pcf":
            add_leads(selPacificAC2)

def add_leads(selopt):
    selector = selopt
    x = True
    while x:
        #print(agentName, tsr, user_campaign, agent_shift)
        AssignUser(input("Enter Agent Name: "))
        input("Enter To Continue")
        SelectQueue(selector)
        SelectLeads(selector)

def SelectQueue(region):
    eastern = "Cell - Eastern"
    central = "Cell - Central"
    pacific = "Cell - Pacific"
    eastern2 = "Cell 2 - Eastern"
    central2 = "Cell 2 - Central"
    pacific2 = "Cell 2 - Pacific"

    #00B38000007XnvD_paginator_rpp_target
    if region == "00B38000007Xnw2_paginator_rpp_target":
        LocOption = eastern
    elif region == "00B38000007Xnvr_paginator_rpp_target":
        LocOption = central
    elif region == "00B38000007XnwQ_paginator_rpp_target":
        LocOption = pacific
    elif region == "00B38000007Xnw1_paginator_rpp_target":
        LocOption = eastern2
    elif region == "00B38000007Xnvw_paginator_rpp_target":
        LocOption = central2
    elif region == "00B38000007XnwV_paginator_rpp_target":
        LocOption = pacific2
    elif region == "00B38000007XnvD_paginator_rpp_target":
        LocOption = eastern
    elif region == "00B38000007Xnv8_paginator_rpp_target":
        LocOption = central
    elif region ==  "00B38000007Xnvh_paginator_rpp_target":
        LocOption = pacific
    elif region ==  "00B38000007XnvD_paginator_rpp_target":
        LocOption = eastern2
    elif region == "00B38000007Xnv8_paginator_rpp_target":
        LocOption = central2
    elif region == "00B38000007Xnvh_paginator_rpp_target":
        LocOption = pacific2

    select = Select(driver.find_element_by_name("fcf"))
    #select.select_by_visible_text("Available - " + LocOption)
    select.select_by_visible_text("Available & Callable - " + LocOption)
    
    time.sleep(4)

def SelectLeads(OptSel):
    print(OptSel)
    by_id(OptSel).click()
    driver.find_element_by_xpath("//*[text()='200']").click()
    time.sleep(1)
    by_id("allBox").click()
    driver.find_element_by_xpath("//input[@type='submit' and @value='Change Owner']").click()

def AssignUser(AgentName):
    time.sleep(3)
    user = by_id('newOwn')
    user.send_keys(AgentName) #Get User Name
    by_name('cancel').click()

def main():
    BrowserSetup()
    startUp()
    SelectLeads(selector)
    leadOpt()

main()
