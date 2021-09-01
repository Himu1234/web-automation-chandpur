from selenium import webdriver
from selenium.webdriver.chrome.options import Options

chrome_options = Options()
# chrome_options.add_argument("--headless")
# #driver = webdriver.Chrome(options=chrome_options)

browser=webdriver.Chrome(executable_path="F:\\web_automation\\chromedriver.exe", options=chrome_options)
browser.get("http://180.211.137.22:8991/Pages/User/BillPrint.aspx")
browser.find_element_by_id("cphMain_txtConsumer")


from selenium.webdriver.common.keys import Keys

x1=browser.find_element_by_id("cphMain_txtConsumer")
x1.send_keys("32673600")
x2=browser.find_element_by_id("cphMain_tbxLocation")
x2.send_keys("B1")
x3=browser.find_element_by_id("cphMain_txtBillCycle")
x3.send_keys("Jan-19")
x3.send_keys(Keys.TAB)
x4=browser.find_element_by_id("cphMain_btnReport")
x4.click()