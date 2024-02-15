import time
import json
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from excel import getreport
from mail import sendmail
from ssologin import ssologin
from tickets import getalltickets

with open('cred.json') as f:
    credentials = json.load(f)

# Use the same user profile as the native Chrome browser
options = webdriver.ChromeOptions()
options.add_argument(r"--start-maximized")
options.add_argument(r"C:\Users\gsingh369\AppData\Local\Google\Chrome\User Data")

# Initialize WebDriver and HPsm plugin
driver = webdriver.Chrome(options=options)
driver.get(credentials["sm9url"])
wait = WebDriverWait(driver, 120)

ssologin(wait, driver, credentials)

time.sleep(20)

# wait for the page title to change
wait.until(EC.title_contains("Service Manager AMS"))

alltickets = getalltickets(wait, driver)                                                

breachcount, excelfile = getreport()

mail_sent = sendmail(excelfile, breachcount)

print("Mail sent = ", mail_sent)
