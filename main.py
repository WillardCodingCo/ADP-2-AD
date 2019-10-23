'''
Author: Alex Willard
Date: 09/23/2019
Description: Syncs ADP data with AD
'''
import os
import re
import sys
import requests
import stat
import uuid
import errno
import shutil
import time
import pyodbc
import subprocess
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

#Global Variables
DOWNLOADPATH = os.getenv('TEMP') + '\\AutoFetch\\Downloads\\'
WEBAPPURL = 'https://script.google.com/macros/s/AKfycbyZnuFTmKm4m3mw50L0A4OjhH3aIZXTQuq1y39InVO2AL3F10I/exec'
DRIVERPATH = '.\\chromedriver.exe'

#Functions
def NewDriver(download_dir):
    options = Options()
    options.add_experimental_option('prefs', {
        'download.default_directory': download_dir,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': True
    })

    driver = webdriver.Chrome(DRIVERPATH, options=options)
    return driver

def startupCheck():
    appdata = os.getenv('TEMP')
    os.makedirs(DOWNLOADPATH, exist_ok=True)
    ClearDownloadDir()
    return True

def ClearDownloadDir():
    file_list = os.listdir(DOWNLOADPATH)
    for f in file_list:
        shutil.rmtree(os.path.join(DOWNLOADPATH, f), ignore_errors=False, onerror=handleRemoveReadonly)

def handleRemoveReadonly(func, path, exc):
  excvalue = exc[1]
  if func in (os.rmdir, os.remove) and excvalue.errno == errno.EACCES:
      os.chmod(path, stat.S_IRWXU| stat.S_IRWXG| stat.S_IRWXO) # 0777
      func(path)
  else:
      raise

def CreateDownloadDir():
    while True:
        r_name = str(uuid.uuid4())
        r_dir = DOWNLOADPATH + r_name
        if os.path.isdir(r_dir):
            continue
        else:
            os.mkdir(r_dir)
            return r_dir

def WaitForDownload(driver):
    driver.get('chrome://downloads/')
    progress = 0
    while progress != 100:
        progress = GetDownloadProgress(driver)

def GetLatestAdpDownloadUrl():
    results = requests.get(WEBAPPURL)
    return results.text

def GetDownloadProgress(driver):
    tries = 0
    while tries < 5:
        try:
            progress = driver.execute_script('''
            
            var tag = document.querySelector('downloads-manager').shadowRoot;
            var intag = tag.querySelector('downloads-item').shadowRoot;
            var progress_tag = intag.getElementById('progress');
            var progress = progress_tag.value
            return progress
            
            ''')
            return progress
        except:
            tries += 1

def GetLatestDownload(download_dir):
    return os.path.join(download_dir,os.listdir(download_dir)[0])

def GetLastWord(str):
    #Handle ()
    if '(' in str:
        str = str[0:str.index('(')]

    text = str.split()
    word = text[-1].replace('?', '')

    return word

#Login handles
def LoginGoogle(driver, email='alex.willard@smdsi.com', password='TeslaR0X'):
    #Don't need in current senerio 
    login_url = 'https://gmail.com'
    driver.get(login_url)
    driver.find_element_by_id('identifierId')
    driver.find_element_by_id('identifierId').send_keys(email)
    driver.find_element_by_id('identifierNext').click()
    time.sleep(2)
    driver.find_element_by_name("password").send_keys(password)
    driver.find_element_by_id("passwordNext").click()
    time.sleep(1)

def LoginGoogleCheck(driver):
    url = driver.current_url

def LoginAdp(driver, username='AWillard@DisasterS', password='bvH3Sx77P4t91vkSfjsm'):
    login_url = 'http://workforcenow.adp.com'
    driver.get(login_url)
    driver.find_element_by_link_text('Administrator Sign In').click()
    time.sleep(3)
    driver.find_elements_by_xpath('//form[@class="adp-form fyp"]//input')[0].send_keys(username)
    driver.find_element_by_xpath('//button[@type="button"]').click()
    time.sleep(2)
    driver.find_elements_by_xpath('//form[@class="adp-form fyp"]//input')[2].send_keys(password)
    driver.find_element_by_xpath('//button[@type="button"]').click()
    time.sleep(1)

def LoginAdpCheck(driver):
    url = driver.current_url

def HandleAdpVerification(driver):
    src = driver.page_source
    text_found = re.search(r'Get a verification code', src)
    #If there is text
    if text_found:
        #Get to questions
         driver.find_element_by_xpath('//span[text()="Answer security questions"]').click()
         driver.find_element_by_css_selector('button.primary.vdl-button.vdl-button--primary').click()

         #Answer Questions
         #Handle ()
         time.sleep(1)
         firstQuestion = driver.find_element_by_xpath('//label[@for="fristQuestion"]').text
         firstAnswer = GetLastWord(firstQuestion)

         secondQuestion = driver.find_element_by_xpath('//label[@for="confirmPassword"]').text
         secondAnswer = GetLastWord(secondQuestion)

         inputs = driver.find_elements_by_xpath('//form[@class="adp-form"]//input[@type="password"]')
         inputs[0].send_keys(firstAnswer)
         inputs[1].send_keys(secondAnswer)

         driver.find_element_by_css_selector('button.primary.vdl-button.vdl-button--primary').click()

    return driver
        



#Main
startupCheck()

#Open Webbrowser
#Set download location
download_dir = CreateDownloadDir()
driver = NewDriver(download_dir)

#Log into ADP and GMail
LoginAdp(driver)
#Handle Security Questions
HandleAdpVerification(driver)
#LoginGoogle(driver)

#Run constant check in https://mail.google.com/mail/u/0/?tab=rm&ogbl#label/Application%2FADP+Report for new mail
url = GetLatestAdpDownloadUrl()

#If new report then click link
driver.get(url) 
#WaitForDownload(driver)
time.sleep(5)
download_file = GetLatestDownload(download_dir)

#Fetch download location and find file 
#print(download_file)

driver.quit()

#Run insert script#
conn = pyodbc.connect('Driver={SQL SERVER};'
	       'Server=Smdsi-WH\SQlEXPRESS;'
	       'Database=Staging;'
	       'Trusted_Connection=yes;')
cursor = conn.cursor()
#Csv 2 Staging.dbo.Adp_ActiveDirectory
cursor.execute('TRUNCATE TABLE Staging.dbo.Adp_ActiveDirectory')
p = subprocess.Popen(["powershell.exe", ".\\import_csv.ps1 '" + download_file + "'"], stdout = sys.stdout)
p.communicate()[0]

#Staging.dbo.Adp_ActiveDirectory 2 Datawarehouse.dbo.Employees
cursor.execute('Exec Staging.dbo.Staging_Adp_ActiveDirectory_To_Datawarehouse_Employees')

conn.close()


#Sql 2 Ad
#p = subprocess.Popen(["powershell.exe", ""])