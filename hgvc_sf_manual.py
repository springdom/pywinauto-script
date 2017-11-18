from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time

user_name = "mtaylor"
passwd = "J@maica1994"

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
    by_name("go").click()
    time.sleep(4)

def leads():
    by_id("00B38000007Xnvh_paginator_rpp_target").click()
    driver.find_element_by_xpath("//*[text()='200']").click()
    time.sleep(1)
    by_id("allBox").click()
    driver.find_element_by_xpath("//input[@type='submit' and @value='Change Owner']").click()

def add_leads():
    x = True
    while x:
        #print(agentName, tsr, user_campaign, agent_shift)
            AssignUser(input("Enter Agent Name: "))
            time.sleep(3)
            SelectQueue()
            time.sleep(2)
            SelectLeads()

def GetUser():
    pass

def SelectQueue():
    select = Select(driver.find_element_by_name("fcf"))
    select.select_by_visible_text("Available & Callable - Cell - Pacific")
    time.sleep(4)

def AssignUser(AgentName):
    time.sleep(3)
    user = by_id('newOwn')
    user.send_keys(AgentName) #Get User Name
    time.sleep(0.3)
    by_name('cancel').click()

def SelectLeads():
    by_id("00B38000007Xnvh_paginator_rpp_target").click()
    driver.find_element_by_xpath("//*[text()='200']").click()
    time.sleep(1)
    by_id("allBox").click()
    driver.find_element_by_xpath("//input[@type='submit' and @value='Change Owner']").click()

def main():
    BrowserSetup()
    startUp()
    leads()
    add_leads()

main()
