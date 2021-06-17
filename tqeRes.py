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
outfile = open("errList", 'w')
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
    vozac.Driver.find_element_by_id("username").send_keys("user.name")
    vozac.Driver.find_element_by_id("password").send_keys("pass")
    vozac.Driver.find_element_by_id("passwordSecondary").send_keys(x)
    vozac.Driver.find_element_by_id("btnSubmit_5").click()
    try:
        vozac.Driver.find_element_by_id("postfixSID_1").click()
        vozac.Driver.find_element_by_id("btnContinue").click()
    except NoSuchElementException:
        return False
    return True

def quitDriver():
    vozac.Driver.quit()

def takeSS():
    vozac.Driver.save_screenshot("screenshot.png")

def clickElem(elemToClick):
    vozac.Driver.find_element_by_xpath(str(elemToClick)).click()
    vozac.Driver.implicitly_wait(2)
    vozac.Driver.save_screenshot("screenshot.png")

def enterTxt(elemToClick, keysToSend):
    vozac.Driver.implicitly_wait(2)
    vozac.Driver.find_element_by_xpath(str(elemToClick)).send_keys(str(keysToSend))

def runFirstTest():
    vozac.Driver.set_page_load_timeout(30)
    vozac.Driver.implicitly_wait(2)

def goToCorpExplain(user, subTaskID):
    per = '%'
    vozac.Driver.get(f'https://scale-dashboard.761link.net/corp/explain?worker={user}{per}40761link.net&subtask={subTaskID}')

def unclaimTask(user, subTaskID):
    vozac.Driver.get(f'https://scale-dashboard.761link.net/corp/lookup/{subTaskID}')
    vozac.Driver.implicitly_wait(2)
    vozac.Driver.find_element_by_xpath('//*[@id="__next"]/div[3]/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[2]/div[1]/div[3]/span/a').click()
    WebDriverWait(vozac.Driver, 10).until(EC.alert_is_present())
    vozac.Driver.switch_to.alert.accept()
    getErrs(user, subTaskID)

def getErrs(user, subTaskID):
    checks = vozac.Driver.find_elements_by_xpath('//*[@id="__next"]/div[3]/div/div[1]/div[2]/div[2]/div/div[1]/div[2]/div')
    errList = []

    for err in checks:
        try:
            hasPerms4FMV = err.find_element_by_xpath('//*[@id="__next"]/div[3]/div/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div[1]/span[3]').text
        except NoSuchElementException:
            hasPerms4FMV = 'N/A'

        subTaskPending = err.find_element_by_xpath('//*[@id="__next"]/div[3]/div/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div[2]/span[3]').text
        stagingBatch = err.find_element_by_xpath('//*[@id="__next"]/div[3]/div/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div[3]/span[3]').text
        inWorkerSkips = err.find_element_by_xpath('//*[@id="__next"]/div[3]/div/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div[4]/span[3]').text
        querySubTaskPerms = err.find_element_by_xpath('//*[@id="__next"]/div[3]/div/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div[5]/span[3]').text
        reviewPermLevel = err.find_element_by_xpath('//*[@id="__next"]/div[3]/div/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div[6]/span[3]').text
        unclaimedOrAssigned = err.find_element_by_xpath('//*[@id="__next"]/div[3]/div/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div[7]/span[3]').text
        excludeWorkers = err.find_element_by_xpath('//*[@id="__next"]/div[3]/div/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div[8]/span[3]').text
        canAccessWithSTID = err.find_element_by_xpath('//*[@id="__next"]/div[3]/div/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div[9]/span[3]').text
        assignedByTrainer = err.find_element_by_xpath('//*[@id="__next"]/div[3]/div/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div[10]/span[3]').text

        if "Subtask is claimed to" in reviewPermLevel:
            claimRes.write(user + ',' + subTaskID + ',' + str(reviewPermLevel) + '\n')
            unclaimTask(user, subTaskID)
        elif "Subtask is claimed to" in unclaimedOrAssigned:
            claimRes.write(user + ',' + subTaskID + ',' + str(unclaimedOrAssigned) + '\n')
            unclaimTask(user, subTaskID)
        else:
            pass

        err_item = {
            'User' : user,
            'Subtask ID' : subTaskID,
            'Has FMV Perms' : hasPerms4FMV,
            'SubTask Pending' : subTaskPending,
            'Staged Batch' : stagingBatch,
            'Skipped Worker' : inWorkerSkips,
            'Can Query' : querySubTaskPerms,
            'Review Perm Level' : reviewPermLevel,
            'Unclaimed/Assigned' : unclaimedOrAssigned,
            'Unexcluded Worker' : excludeWorkers,
            'Can Access with Subtask ID' : canAccessWithSTID,
            'Assigned By Trainer' : assignedByTrainer
        }

        userName = str(err_item.get('User'))
        subID = str(err_item.get('subTaskID'))
        hasPerms = str(err_item.get('Has FMV Perms'))
        taskStatus = str(err_item.get('SubTask Pending'))
        workerSkipped = str(err_item.get('Skipped Worker'))
        rPermLevel = str(err_item.get('Review Perm Level'))
        claimAssign = str(err_item.get('Unclaimed/Assigned'))
        unEx = str(err_item.get('Unexcluded Worker'))

        errList.append(err_item)
        if("Subtask is unclaimed" in reviewPermLevel):
            of2.write(userName + ',' + subTaskID + ',' + hasPerms + ',' + taskStatus + ',' + workerSkipped + ',' + 'NULL' + ',' + rPermLevel + ',' + claimAssign + ',' + unEx + ',\n')
        else:
            of2.write(userName + ',' + subTaskID + ',' + hasPerms + ',' + taskStatus + ',' + workerSkipped + ',' + rPermLevel + ',' + claimAssign + ',' + unEx + ',\n')
    df = pd.DataFrame(errList)
    df.to_csv('tqeResolve.csv', mode = 'a')

    print(df)

def tqeList():
    with open("kMetrics - rSheet.csv", 'r') as infile:
        reader = csv.reader(infile, delimiter=",")

        for row in reader:
            subTaskID = row[0]
            user = row[1]
            goToCorpExplain(user, subTaskID)
            vozac.Driver.implicitly_wait(2)
            getErrs(user, subTaskID)
            outfile.write("{}, {}\n".format(subTaskID, user))
    outfile.close()
            
def getVideoDetail():
    videos = vozac.Driver.find_elements_by_class_name('style-scope ytd-grid-video-renderer')

    vid_list = []

    for video in videos:
        title = video.find_element_by_xpath('.//*[@id="video-title"]').text
        views = video.find_element_by_xpath('.//*[@id="metadata-line"]/span[1]').text
        uploadTS = video.find_element_by_xpath('.//*[@id="metadata-line"]/span[2]').text
        vid_item = {
            'title' : title,
            'views' : views,
            'posted': uploadTS
        }

        vid_list.append(vid_item)
    df = pd.DataFrame(vid_list)
        
    print(df)

def sendToSheet():
    content = open('TQEs.csv', 'r').read()
    clients.import_csv('1f33kclXV4MQk_XXXXX', data=content)

def mainQ():
    goToCorp()
    loginToCorp()
    uIn = input('Continue?')
    print(uIn)
    tqeList()
    of2.close()
    sendToSheet()
    uIn = input('Quit?')
    vozac.Driver.quit()

mainQ()
