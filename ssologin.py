import pyautogui
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC


def ssologin(wait, driver, credentials):
    
    # Use WebDriverWait to wait for the element to be present before trying to interact with it
    email_field_xpath = '/html/body/div[2]/main/div[2]/div/div/div[2]/form/div[1]/div[3]/div/div[2]/span/input'
    email_field = wait.until(EC.presence_of_element_located((By.XPATH, email_field_xpath)))
    email_field.send_keys(credentials["mailid"])
    email_field.send_keys(Keys.RETURN)

    # Use WebDriverWait to wait for the element to be present before trying to interact with it
    ssopass_field_xpath = '/html/body/div[2]/main/div[2]/div/div/div[2]/form/div[1]/div[4]/div/div[2]/span/input'
    ssopass_field = wait.until(EC.presence_of_element_located((By.XPATH, ssopass_field_xpath)))
    ssopass_field.send_keys(credentials["ssopass"])
    ssopass_field.send_keys(Keys.RETURN)

    # Wait for the button to be present on the next page
    button_xpath = '/html/body/div[2]/main/div[2]/div/div/div[2]/form/div[2]/div/div[3]/div[2]/div[2]/a'
    wait.until(EC.presence_of_element_located((By.XPATH, button_xpath)))
    button = driver.find_element(By.XPATH, button_xpath)
    button.click()

    time.sleep(2)
    pyautogui.press("tab")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(2)

    pyautogui.press("tab")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)

    pyautogui.press("tab")
    time.sleep(1)
    pyautogui.press("tab")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.typewrite(credentials["pin"], interval=0.1)
    time.sleep(1)
    pyautogui.press("enter")
