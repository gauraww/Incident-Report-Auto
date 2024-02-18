from json import load
from sys import exit
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from excel import getreport
from mail import sendmail
from ssologin import ssologin
from tickets import getalltickets

with open('cred.json') as f:
    credentials = load(f)

# Use the same user profile as the native Chrome browser
options = webdriver.ChromeOptions()
options.add_argument(r"--start-maximized")
options.add_argument(r"C:\Users\gsingh369\AppData\Local\Google\Chrome\User Data")

# Initialize WebDriver and HPsm plugin
driver = webdriver.Chrome(options=options)
driver.get(credentials["sm9url"])
wait = WebDriverWait(driver, 30)

ssologin(wait, driver, credentials)

# Wait for the page title to change or refresh if it takes more than 30 seconds
try:
    wait.until(EC.title_contains("Service Manager AMS"))
except TimeoutException:
    driver.get(credentials["sm9url"])
    wait.until(EC.title_contains("Service Manager AMS"))

getalltickets(wait, driver)                                                

breachcount, excelfile = getreport()

sendmail(excelfile, breachcount)

exit()