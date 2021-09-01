import os
from selenium import webdriver
from time import sleep
from openpyxl import load_workbook
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.common.action_chains import ActionChains

driver_exe = 'chromedriver.exe'
driver = webdriver.Chrome(executable_path=os.path.join(os.getcwd(), driver_exe))
#driver.maximize_window()
driver.get("https://www.autodraw.com/")
driver.find_element_by_css_selector(".buttons > .green").click()
driver.find_element_by_xpath('/html/body/div/div[2]/div[3]/div[3]/img').click()
canvas = driver.find_element_by_xpath('/html/body/div/div[2]/div[4]/canvas')

action = ActionChains(driver)
action.move_to_element(canvas)
action.click_and_hold()
action.move_by_offset(8,1)
action.move_by_offset(6,1)
action.move_by_offset(4,1)
action.move_by_offset(2,1)
action.move_by_offset(1,1)
action.move_by_offset(1,2)
action.move_by_offset(1,4)
action.move_by_offset(1,6)
action.move_by_offset(1,8)
action.release()
sleep(5)
print('Sleep Ends')
action.perform()

x = driver.find_element_by_xpath('/html/body/div/div[2]/div[3]/div[4]/img')
x.click()
action.move_to_element(canvas)
action.move_by_offset(1,8)
action.click()
sleep(5)
print('Sleep Ends - 2')
action.perform()
sleep(4)
driver.close()


