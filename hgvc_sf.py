from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
import time

user_name = ""
passwd = ""

driver = webdriver.Firefox()
driver.get("https://hgv.my.salesforce.com")
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
time.sleep(6)
by_id("00B38000007XnvD_paginator_rpp_target").click()
#elem.send_keys(Keys.RETURN)
#driver.close()

