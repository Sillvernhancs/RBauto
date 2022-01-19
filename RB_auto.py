from email import message
from operator import add
from sys import exc_info
from tracemalloc import start
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common import by
from selenium.common.exceptions import NoSuchElementException  
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import re
import time
#///////////// safe link stuff
import posixpath as path
from urllib.parse import urlparse, parse_qs, urlunparse
#/////////////////////////////////////////////////////////////////////
import win32com.client
# /////////////////////////////////////////////////////////////////////
# close all tabs
def closeAllTabs(brwsr):
    for handle in brwsr.window_handles:
        brwsr.switch_to.window(handle)
        brwsr.close()
def init_browser(start_URL):
    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    service_ = Service("chromedriver.exe")
    b = webdriver.Chrome(service=service_, options=chrome_options)

    b.set_window_size(1080,1080)
    b.get(start_URL)
    #returns a browser obj
    return b
def login(netID, password):
    print('Trying to login....')
    print('___________________')
    browser = init_browser('https://groups.accc.uic.edu/login')
    browser.find_element(By.ID,'inputUserid').send_keys(netID)
    browser.find_element(By.ID,'inputPassword').send_keys(password)
    browser.find_element(By.ID,"login-button").click()
    try:
        WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.XPATH,"/html/body/div[3]/div/div[1]/div[1]/div/div[2]/a"))
        ).click()
        closeAllTabs(browser)
        return True
    except:
        closeAllTabs(browser)
        return False

def addNetID(NID, URL):
    browser = init_browser('https://groups.accc.uic.edu/login')
    browser.find_element(By.ID,'inputUserid').send_keys(netID)
    browser.find_element(By.ID,'inputPassword').send_keys(password)
    browser.find_element(By.ID,"login-button").click()
    # click on VPN button
    WebDriverWait(browser, 3).until(
    EC.presence_of_element_located((By.XPATH,"/html/body/div[3]/div/div[1]/div[1]/div/div[2]/a"))
    ).click()
    # COP- affiliates
    # browser.find_element(By.XPATH,"/html/body/div[3]/div/div/div/table/tbody[2]/tr[31]/td[4]/span/form/button[2]/svg").click()
    element = browser.find_element(By.XPATH,'//*[@id="app"]/div/div/div/table/tbody[2]/tr[31]/td[4]/span/form/button[2]')

    # actions = ActionChains(browser)
    # actions.move_to_element(element).perform()
    element.click()
    # Add member
    browser.find_element(By.XPATH,"/html/body/div[3]/div[1]/div[3]/div/div/span/button").click()
    # fill in info
    time.sleep(1.5)
    WebDriverWait(browser, 3).until(
    EC.presence_of_element_located((By.XPATH,'//*[@id="add-netid"]'))
    ).send_keys(NID)
    WebDriverWait(browser, 3).until(
    EC.presence_of_element_located((By.XPATH,'//*[@id="add-rationale"]'))
    ).send_keys(URL)
    browser.find_element(By.XPATH,'//*[@id="addForm"]').click()
    closeAllTabs(browser)
    return
#/////////////////////////////////////////////////////////////////////
# Main:... 
while True:
    # get user credentials
    print("/////////////////////////////////")
    netID    = input("NetID   : ")
    password = input("Password: ")
    print("/////////////////////////////////")
    if(login(netID,password)):
        print('Login successful')
        break
    else:
        print('Login failed, try again...')
        continue
print('Waiting for a Remote Access Request Form...')
while True:
    ol = win32com.client.Dispatch( "Outlook.Application")
    inbox = ol.GetNamespace("MAPI").GetDefaultFolder(6)
    messages = inbox.Items
    # get the last item in the inbox
    message_current = messages.GetLast()
    # scan 10 items from the last email received (increase or decrease if need be)
    for i in range(10):
        if (message_current.UnRead == True) and ('You are receiving this email from Qualtrics in response to a Remote Access Request from the user below. ' in message_current.Body):
            # set as read to avoid dupications.
            message_current.UnRead = False
            # get netID
            NETID = message_current.Body[message_current.Body.find('UIN)') + 9:message_current.Body.find('Phone Number or extension') - 3]
            #grab link (safe link form need to be decoded)
            URL = re.search("(?P<url>https?://[^\s]+)", message_current.Body).group("url")
            print ('Adding user: ' + NETID)
            target = parse_qs(urlparse(URL).query)['url'][0]
            p = urlparse(target)
            q = p._replace(path=path.join(path.dirname(path.dirname(p.path)), path.basename(p.path)))
            #decoded safelinks
            URL =  (urlunparse(q))
            print ('With URL: ' + URL)
            print ('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
            addNetID(NETID, URL)
        # go to the next email
        message_current = messages.GetPrevious()
    # delay of 5s for every scans.
    time.sleep(5)