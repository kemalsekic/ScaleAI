import gspread
from oauth2client.service_account import ServiceAccountCredentials
import vozac
import pandas as pd
import csv
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive","https://www.googleapis.com/auth/spreadsheets.readonly"]
creds = ServiceAccountCredentials.from_json_keyfile_name("kCreds.json", scope)
clients = gspread.authorize(creds)
spreadsheet = clients.open("kBot")
sheet = spreadsheet.get_worksheet(3)
secondSheet = spreadsheet.sheet1

url = 'https://scale-dashboard.net/corp/login?redirect_url=XXXXXXXX'
outfile = open("officialReport.csv", 'w')
of2 = open("TQEs.csv", 'w')
of2_header = "User, Subtask ID, Has Permissions, Subtask Status, Skipped Worker, Review Permission Level, Claimed, Excluded/Assignment\n"
of2.write(of2_header)
claimRes = open("claimRes.csv", 'w')
claimRes.write("User, Subtask ID, Claimant\n")

def goToCorp():
    vozac.Driver.set_page_load_timeout(30)
    vozac.Driver.implicitly_wait(2)
    vozac.Driver.get(url)

def check_exists_by_xpath(xpath):
    try:
        vozac.Driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

def loginToCorp():
    x = input("Secondary Pass: ")
    print("Thank you")
    clickElem("/html/body/div[2]/div[3]/div/div[3]/button/span[1]")
    clickElem("/html/body/div/div[3]/div[3]/div/div/a/span")
    vozac.Driver.find_element_by_id("username").send_keys("user.name") #Add your cerdentials here
    vozac.Driver.find_element_by_id("password").send_keys("pass") #Add your cerdentials here
    vozac.Driver.find_element_by_id("passwordSecondary").send_keys(x) #Add your secondary passcode from the swivel app here
    vozac.Driver.find_element_by_id("btnSubmit_5").click()
    try:
        vozac.Driver.find_element_by_id("postfixSID_1").click()
        vozac.Driver.find_element_by_id("btnContinue").click()
    except NoSuchElementException:
        return False
    return True

def goToCorpExplain(subTaskID):
    vozac.Driver.get(f'https://scale-dashboard.net/corp/lookup/{subTaskID}')
    try:
        status = vozac.Driver.find_element_by_xpath('//*[@id="__next"]/div[3]/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[1]/div/div/div[5]/span[2]').text
        print(str(status))
    except NoSuchElementException:
        return False
    return str(status)

def tqeList():
    with open("taskList.csv", 'r') as infile:
        reader = csv.reader(infile, delimiter=",")

        for row in reader:
            subTaskID = row[0]
            stat = goToCorpExplain(subTaskID)
            vozac.Driver.implicitly_wait(2)
            outfile.write("{}, {}\n".format(subTaskID, str(stat)))
    outfile.close()

def quitDriver():
    vozac.Driver.quit()

def clickElem(elemToClick):
    vozac.Driver.find_element_by_xpath(str(elemToClick)).click()
    vozac.Driver.implicitly_wait(2)

def enterTxt(elemToClick, keysToSend):
    vozac.Driver.implicitly_wait(2)
    vozac.Driver.find_element_by_xpath(str(elemToClick)).send_keys(str(keysToSend))

def sendToSheet():
    content = open('officialReport.csv', 'r').read()
    clients.import_csv('1f33kclXV4MQk_XXXXX', data=content)

def mainQ():
    goToCorp()
    loginToCorp()
    tqeList()
    sendToSheet()
    uIn = input('Quit?')
    vozac.Driver.quit()

mainQ()
