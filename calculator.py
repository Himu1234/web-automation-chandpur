from selenium import webdriver

driver = webdriver.Remote(
    command_executor='http://localhost:9999',
    desired_capabilities={
        "debugConnectToRunningApp": 'false',
        "app": r"C:/Windows/System32/calc.exe"
    })

window = driver.find_element_by_class_name('CalcFrame')
view_menu_item = window.find_element_by_id('MenuBar').find_element_by_name('View')

view_menu_item.click()
view_menu_item.find_element_by_name('Scientific').click()

view_menu_item.click()
view_menu_item.find_element_by_name('History').click()

window.find_element_by_id('132').click()
window.find_element_by_id('93').click()
window.find_element_by_id('134').click()
window.find_element_by_id('97').click()
window.find_element_by_id('138').click()
window.find_element_by_id('121').click()

driver.close()