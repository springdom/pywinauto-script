from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time

driver = webdriver.Firefox()
by_name = driver.find_element_by_name
by_id = driver.find_element_by_id

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
    SelectQueue()
    try:
        by_name("go").click()
    except:
        pass
    
    time.sleep(4)

selector =    "00B38000007Xnw2_paginator_rpp_target"

selPacific =  "00B38000007XnwQ_paginator_rpp_target"
selPacific2 = "00B38000007XnwV_paginator_rpp_target"
selPacificAC = "00B38000007Xnvh_paginator_rpp_target"
#selPacificAC2 = "00B38000007Xnvh_paginator_rpp_target"

selEastern =  "00B38000007Xnw2_paginator_rpp_target"
selEastern2 = "00B38000007Xnw1_paginator_rpp_target"
selEasternAC = "00B38000007XnvD_paginator_rpp_target"
#selEasternAC2 = "00B38000007XnvD_paginator_rpp_target"

selCentral =  "00B38000007Xnvr_paginator_rpp_target"
selCentral2 = "00B38000007Xnvw_paginator_rpp_target"
selCentralAC = "00B38000007Xnv8_paginator_rpp_target"
#selCentralAC2 = "00B38000007Xnv8_paginator_rpp_target"



def add_leads():
    x = True
    while x:
        #print(agentName, tsr, user_campaign, agent_shift)
            AssignUser(input("Enter Agent Name: "))
            input("Enter To Continue")
            SelectQueue()
            SelectLeads(selector)

def SelectQueue():
    eastern = "Cell - Eastern"
    eastern2 = "Cell 2 - Eastern"
    pacific = "Cell - Pacific"
    pacific2 = "Cell 2 - Pacific"
    central = "Cell - Central"
    central2 = "Cell 2 - Central"
    
    select = Select(driver.find_element_by_name("fcf"))
    
    select.select_by_visible_text("Available - " + eastern)
    #select.select_by_visible_text("Available & Callable - " + pacific)
    
    time.sleep(4)

def SelectLeads(Selector):
    by_id(Selector).click()
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
    add_leads()

main()
