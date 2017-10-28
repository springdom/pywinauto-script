from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time

user_name = "mtaylor@hgvc.com"
passwd = "J@maica1994"

driver = webdriver.Firefox()
driver.get("https://hgv.my.salesforce.com/home/home.jsp?tsid=02u38000000NcOF")
assert "Salesforce" in driver.title
by_name = driver.find_element_by_name
by_id = driver.find_element_by_id

username = by_name('username')
username.send_keys(user_name)
password = by_name('pw')
password.send_keys(passwd)

by_id("Login").click()
time.sleep(6)
by_id("01r38000000Hv1x_Tab").click()
time.sleep(2)

select = Select(driver.find_element_by_id("fcf"))
select.select_by_visible_text("Available & Callable - Cell - Pacific")
by_name("go").click()
time.sleep(4)

#Loop Here?
by_id("00B38000007Xnvh_paginator_rpp_target").click()
driver.find_element_by_xpath("//*[text()='200']").click()

time.sleep(1)
by_id("allBox").click()
driver.find_element_by_xpath("//input[@type='submit' and @value='Change Owner']").click()
#by_id("fcf").click()

#elem.send_keys(Keys.RETURN)
#driver.close()


