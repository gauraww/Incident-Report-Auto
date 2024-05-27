from json import load
from sys import exit
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from excel import getreport
from mail import sendmail
from ssologin import ssologin
from tickets import getalltickets

with open('cred.json') as f:
    credentials = load(f)

# Initialize WebDriver and set Chrome options
chrome_service = Service(ChromeDriverManager().install())
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--start-maximized")

# Create WebDriver instance
driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

# Navigate to the URL
driver.get(credentials["sm9url"])

# Initialize WebDriverWait
wait = WebDriverWait(driver, 30)

# Wait for the page title to change or refresh if it takes more than 30 seconds
try:
    ssologin(wait, driver, credentials)
    wait.until(EC.title_contains("Service Manager AMS"))
except TimeoutException:
    driver.refresh()
    wait.until(EC.title_contains("Service Manager AMS"))

getalltickets(wait, driver)     

breachcount, excelfile = getreport()

sendmail(excelfile, breachcount)
