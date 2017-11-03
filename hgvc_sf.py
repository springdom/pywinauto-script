from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from openpyxl import load_workbook
import time

user_name = "mtaylor"
passwd = "J@maica1994"

driver = webdriver.Firefox()
driver.get("https://identitynow.hgv.com/")
#driver.get("https://hgv.my.salesforce.com/home/home.jsp?tsid=02u38000000NcOF")
#assert "HGV IdentityNow" in driver.title
by_name = driver.find_element_by_name
by_id = driver.find_element_by_id


username = by_id('username')
username.send_keys(user_name)
password = by_id('password')
password.send_keys(passwd)


by_id("signIn").click()
time.sleep(5)
driver.get("https://hgv.my.salesforce.com/home/home.jsp?tsid=02u38000000NcOF")
time.sleep(7)
by_id("01r38000000Hv1x_Tab").click()
time.sleep(2)

select = Select(driver.find_element_by_id("fcf"))
select.select_by_visible_text("Available & Callable - Cell - Pacific")
by_name("go").click()
time.sleep(4)

#Loop Here?
by_id("00B38000007Xnvh_paginator_rpp_target").click()
driver.find_element_by_xpath("//*[text()='200']").click()

# Or hereLoop Here?
time.sleep(1)
by_id("allBox").click()
driver.find_element_by_xpath("//input[@type='submit' and @value='Change Owner']").click()

user = by_id('newOwn')
user.send_keys('Test User')
#driver.find_element_by_xpath("//input[@type='submit' and @value='Save']").click()

#driver.close()



