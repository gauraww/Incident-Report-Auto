from datetime import datetime,timedelta
from time import sleep
from pyautogui import press, hotkey, typewrite, write, click
from pyperclip import paste
from selenium.webdriver.common.by import By   
from selenium.webdriver.support import expected_conditions as EC


def getalltickets(wait, driver):
    # press incident mgmt button
    button_xpath = '/html/body/div[2]/div/div[2]/div[1]/div/div/div[4]/div[1]/span'
    wait.until(EC.presence_of_element_located((By.XPATH, button_xpath)))
    button = driver.find_element(By.XPATH, button_xpath)
    button.click()

    # press incident search button
    button_xpath = '/html/body/div[2]/div/div[2]/div[1]/div/div/div[4]/div[2]/div/ul/div/li[3]/div/a'
    wait.until(EC.presence_of_element_located((By.XPATH, button_xpath)))
    button = driver.find_element(By.XPATH, button_xpath)
    button.click()

    # search page check
    search_xpath = '/html/body/div[3]/div[2]/div/div[2]/div[1]/div/div[5]/div/div[1]/div/table/tbody/tr/td[1]'
    wait.until(EC.presence_of_element_located((By.XPATH, search_xpath)))

    sleep(10)

    for _ in range(33):
        press('tab')
        sleep(0.1)

    press('right')
    sleep(0.5)
    press('right')

    for _ in range(3):
        press('tab')
        sleep(0.1)

    sleep(0.5)
    typewrite('W-INCFLS-VPC-STORAGE')
    sleep(0.5)

    for _ in range(14):
        press('tab')
        sleep(0.1)

    sleep(0.5)
    hotkey('shift', 'enter')

    sleep(0.5)
    for _ in range(2):
        press('tab')
        sleep(0.1)

    sleep(0.5)
    hotkey('shift', 'enter')

    sleep(2)

    hotkey('ctrl', 'c')
    sleep(1)
    
    datetime_str = paste()
    sleep(1)
    datetime_obj = datetime.strptime(datetime_str, "%d/%m/%Y %H:%M:%S")
    new_datetime_obj = datetime_obj - timedelta(days=1)
    new_datetime_obj = new_datetime_obj.replace(second=0)   
    new_datetime_str = new_datetime_obj.strftime("%d/%m/%Y %H:%M:%S")
    sleep(1)
    write(new_datetime_str)
    sleep(2)
    press('enter')

    sleep(20)

    click(x=1845, y=295)

    # press('tab', presses=95, interval=0.1)
    # press('enter')

    sleep(2)
    hotkey('win', 'up')
    sleep(1)
    hotkey('ctrl', 'a')
    hotkey('ctrl', 'c')

    sleep(2)
    # hotkey('alt', 'f4')
    sleep(1)
    driver.quit()

