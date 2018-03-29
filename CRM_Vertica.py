# Script Name: CRM_Verticas.py
# Date: 02/15/2018 (February 15, 2018)
# Description: This script will run every Monday or when requested, it will access the log in page for infinity.infinityauto.com
# and search for specific .xml files with the "REimport" name inside CRM and will download the REimport view files from 0-12 months ago.
# It will also automatically close after 3 minutes after all tasks have been executed.
# Developed by Sergio A. Nunez and Joel Gutierrez
# Collaborators: Amanda De Leon

import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

print("Starting process...")

# Set up Chrome for disabling download warnings
opts = webdriver.ChromeOptions()
opts.add_experimental_option('prefs', {'safebrowsing.enabled': True})
opts.add_argument('--safebrowsing-disable-download-protection')

# Set up Chrome with the disable properties
driver = webdriver.Chrome(chrome_options=opts)
driver.maximize_window()
driver.set_page_load_timeout(30)
driver.get("https://infinity.infinityauto.com")

# Log In with username and password
login = driver.find_element_by_xpath("//input[@id='userNameInput']")
login.clear()
login.send_keys("ad\deleona" + Keys.TAB + "Infinity20" + Keys.ENTER)
driver.set_page_load_timeout(30)
time.sleep(5)
print("Logged in")

# Close the tutorial
driver.switch_to.frame("InlineDialog_Iframe")
driver.find_element_by_xpath("//*[@class='navTourClose']").click()
driver.switch_to.default_content()
print("Tutorial Closed")
time.sleep(5)

# Browse through Sales and Policy Quotes
driver.find_element_by_css_selector("a.navTabButtonLink > span.navTabButtonImageContainer").click()
time.sleep(5)
driver.find_element_by_id("SFA").click()
time.sleep(3)
print("Sales > ", end="")
driver.find_element_by_id("TabSFA").click()
time.sleep(10)
driver.find_element_by_id("rightNavLink").click()
time.sleep(10)
driver.find_element_by_id("netics_policyquote").click()
time.sleep(5)
print("Policy Quotes")

# Get the list with reported data
# This part involved creating two separate lists, one for the original list and one with a duplicated list,
# the duplicated list is the one where we'll be storing our REimport views to download.
driver.switch_to.frame("contentIFrame0")
time.sleep(5)
driver.find_element_by_xpath("//*[@id='crmGrid_SavedNewQuerySelector']/span[2]").click()
time.sleep(5)
list = driver.find_element_by_xpath("//ul[@class='ms-crm-VS-Menu']")
sublist = list.find_elements_by_tag_name("li")
reImports = []
for index in sublist:
    if "REimport" in index.text:
        reImports.append(index)

# Each REimport view has an index so I iterated through each index to download the view.
# This contained multiple HTML frames to switch from to download the requested views.
for index in reImports:
    print("%s" % index.text, end=" ")
    index.click()
    driver.switch_to.default_content()
    driver.find_element_by_xpath("//*[@id='netics_policyquote|NoRelationship|HomePageGrid|Mscrm.HomepageGrid.netics_policyquote.ExportToExcel']/span/a").click()
    time.sleep(4)
    driver.switch_to.frame("InlineDialog_Iframe")
    driver.find_element_by_xpath("//*[@id='chkReimport']").click()
    driver.find_element_by_xpath("//*[@id='printAll']").click()
    driver.find_element_by_xpath("//*[@id='dialogOkButton']").click()
    time.sleep(10)
    print("Downloaded")
    driver.switch_to.frame("contentIFrame0")
    driver.find_element_by_xpath("//*[@id='crmGrid_SavedNewQuerySelector']/span[2]").click()

# After finishing the view downloads, the program will have a 3 minute pause afterwards to close and quit the driver.
time.sleep(180)
driver.close()
driver.quit()

print("DONE!")
