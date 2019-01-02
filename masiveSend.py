# Note: For proper working of this Script Good and Uninterepted Internet Connection is Required
# Keep all contacts unique
# Can save contact with their phone Number

# Import required packages
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

import datetime
import time
import openpyxl as excel



# function to read contacts from a text file
def readContacts(fileName):
    lst = []
    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['A']

    for cell in range(len(firstCol)):
        contact = str(firstCol[cell].value)
        lst.append(contact)
    return lst

# Target Contacts, keep them in double colons
# Not tested on Broadcast
targets = readContacts("contacts.xlsx")

# can comment out below line
print(targets)


chrome_options = Options()
chrome_options.add_argument('lang=en')
# chrome_options.add_argument("headless")
chrome_options.add_argument("disable-gpu")
chrome_options.add_argument("no-sandbox")
chrome_options.add_experimental_option('prefs', {
    "protocol_handler.excluded_schemes":{
    "afp":True,
    "data":True,
    "disk":True,
    "disks":True,
    "file":True,
    "hcp":True,
    "intent":True,
    "itms-appss":True,
    "itms-apps":True,
    "itms":True,
    "market":True,
    "javascript":True,
    "mailto":True,
    "ms-help":True,
    "news":True,
    "nntp":True,
    "shell":True,
    "sip":True,
    "snews":False,
    "vbscript":True,
    "view-source":True,
    "vnd":{
        "ms":{
            "radio":True
            }
        }
    }
    }
)


# Driver to open a browser
driver = webdriver.Chrome(chrome_options=chrome_options)


def super_get(url):
    try:
        driver.get(url)
        driver.execute_script("window.alert = function() {};")
    except Exception as e:
        alert = driver.switch_to_alert()
        alert.accept()

# Open whatsapp
driver.get("https://web.whatsapp.com/")
wait = WebDriverWait(driver, 10)
wait5 = WebDriverWait(driver, 5)
input("Scan the QR code and then press Enter")


for phone in targets:

    url = "https://api.whatsapp.com/send?phone="+str(phone)+"&text=HolaMundo-TestArtool-Masivo"
    print(url)
    #link to open a site
    super_get(url)


    wait = WebDriverWait(driver, 10)
    wait5 = WebDriverWait(driver, 5)

    x_arg = '//*[@id="action-button"]'
    driver.find_element_by_xpath(x_arg).click()
    print("click in button")


    wait = WebDriverWait(driver, 10)
    wait5 = WebDriverWait(driver, 5)


    # Select the Input Box
    inp_xpath = "//div[@contenteditable='true']"
    input_box = wait.until(EC.presence_of_element_located((By.XPATH, inp_xpath)))
    time.sleep(1)

    # Send message in whatsapp
    input_box.send_keys(Keys.ENTER)
    print("Successfully Send Message")

    time.sleep(0.5)


input("Stop")

driver.quit()
