import os
import re
from selenium import webdriver
from time import sleep
from openpyxl import load_workbook
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.support.ui import WebDriverWait # Required for explicit wait
from selenium.webdriver.support import expected_conditions as ec # Required for explicit wait
from selenium.webdriver.common.by import By # Required for explicit wait

driver_exe = 'chromedriver.exe'
browser = webdriver.Chrome(executable_path=os.path.join(os.getcwd(), driver_exe))

browser.get('http://119.40.95.163/Admin')
browser.implicitly_wait(100)
browser.find_element_by_xpath('/html/body/div/login/form/md-card/md-card-content/md-input-container[1]/input').send_keys('ae1c6')
browser.find_element_by_xpath('/html/body/div/login/form/md-card/md-card-content/md-input-container[2]/input').send_keys('c6_new_connection')
browser.find_element_by_xpath('/html/body/div/login/form/md-card/md-card-actions/button[1]').click()
browser.implicitly_wait(100)
browser.maximize_window()
sleep(2)

for x in range(500):

    browser.get('http://119.40.95.163/Census/CensusVerifyByAe')
    browser.implicitly_wait(100)
    browser.find_element_by_xpath('/html/body/section[2]/div/div/trackingserial/div/div/md-input-container[1]/md-autocomplete/md-autocomplete-wrap/input').click()
    browser.find_element_by_xpath('/html/body/md-virtual-repeat-container/div/div[2]/ul/li[1]/md-autocomplete-parent-scope/span').click()
    browser.find_element_by_xpath('/html/body/section[2]/div/div/trackingserial/div/div/md-input-container[2]/md-select').click()
    browser.find_element_by_xpath('/html/body/div[6]/md-select-menu/md-content/md-option').click()  
    browser.implicitly_wait(100)
    browser.find_element_by_xpath('/html/body/section[2]/div/div/div/div/md-radio-group/md-radio-button[1]').click()
    browser.implicitly_wait(100)
    sleep(1)
    browser.find_element_by_xpath('/html/body/section[2]/div/div/div/button/span').click()
    browser.implicitly_wait(100)
    browser.find_element_by_xpath('/html/body/div[6]/div[7]/div/button').click()
    browser.implicitly_wait(500)
    print('End: ', x)
    
browser.close()