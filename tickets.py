import pyautogui
from datetime import datetime,timedelta
import time
import pyautogui
import pyperclip
import time
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

    time.sleep(10)

    for _ in range(33):
        pyautogui.press('tab')
        time.sleep(0.1)

    pyautogui.press('right')
    time.sleep(0.5)
    pyautogui.press('right')

    for _ in range(3):
        pyautogui.press('tab')
        time.sleep(0.1)

    time.sleep(1)
    pyautogui.typewrite('W-INCFLS-VPC-STORAGE')
    time.sleep(1)

    for _ in range(14):
        pyautogui.press('tab')
        time.sleep(0.1)

    time.sleep(1)
    pyautogui.hotkey('shift', 'enter')

    time.sleep(1)
    for _ in range(2):
        pyautogui.press('tab')
        time.sleep(0.1)

    time.sleep(1)
    pyautogui.hotkey('shift', 'enter')

    time.sleep(3)

    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)
    
    datetime_str = pyperclip.paste()
    print(datetime_str)
    time.sleep(1)
    datetime_obj = datetime.strptime(datetime_str, "%d/%m/%Y %H:%M:%S")
    new_datetime_obj = datetime_obj - timedelta(days=1)
    new_datetime_obj = new_datetime_obj.replace(second=0)
    new_datetime_str = new_datetime_obj.strftime("%d/%m/%Y %H:%M:%S")

    print(new_datetime_str)
    time.sleep(1)
    pyautogui.write(new_datetime_str)
    time.sleep(1)
    pyautogui.press('enter')

    time.sleep(15)

    pyautogui.press('tab', presses=91, interval=0.1)
    pyautogui.press('enter')

    time.sleep(10)
    pyautogui.hotkey('win', 'up')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')

    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')

