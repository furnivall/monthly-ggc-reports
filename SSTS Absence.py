'''
Imagine you are coming back to this file in a year, or you are handing
it over to someone else. Including a little docstring like this one
will help. 
'''
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import os
import configparser
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
import pandas as pd
from pandas.tseries.offsets import MonthEnd
from dateutil.relativedelta import *  # Why not specify these as well?

finperiods = {'04-19': [1, 1, 4],
              '05-19': [2, 5, 8],
              '06-19': [3, 9, 13],
              '07-19': [4, 14, 17],
              '08-19': [5, 18, 21],
              '09-19': [6, 22, 26],
              '10-19': [7, 27, 30],
              '11-19': [8, 31, 34],
              '12-19': [9, 35, 39],
              '01-20': [10, 40, 43],
              '02-20': [11, 44, 47],
              '03-20': [12, 48, 52]}

path = "W:/Workforce Information/Database/Absence/Absence_Working_Files/"
path2 = "W:/Workforce Information/Database/Employee_Leavers/Employee_Working_Files"
date = input("Which month is the target month? (format = MM/YYYY)")
print(date)
date = pd.to_datetime(date)
finweeks = (list(range(finperiods[(date.strftime('%m-%y'))][1], finperiods[(date.strftime('%m-%y'))][2]+1)))
finweeks = ([str(date.year)+'W'+f'{i:02}' for i in finweeks])
finmonth = str(date.year)+'M'+f'{finperiods[(date.strftime("%m-%y"))][0]:02}'
enddate = date + MonthEnd(1)
sickdate = date - relativedelta(months=2)
leavedate = date - relativedelta(months=1)
leaveenddate = date+MonthEnd(1)
wstats18 = date - relativedelta(months=18) + MonthEnd(1)

print(enddate)

config = configparser.ConfigParser()
config.read(r'W:\\Python\Danny\SSTS Extract\SSTSConf.ini')

chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory": r"W:\Workforce Information\Database\Absence\Absence_Working_Files",
         'safebrowsing.disable_download_protection': True}
chromeOptions.add_experimental_option("prefs", prefs)
browser = webdriver.Chrome(executable_path="W:/Danny/Chrome Webdriver/chromedriver.exe",
                           options=chromeOptions)
actionChains = ActionChains(browser)
filename = "W:/Workforce Information/Database/Absence/Absence_Working_Files/Marion-Absence.xls"


def login():
    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/logon.jsp')
    time.sleep(2)
    browser.switch_to.frame('infoView_home')
    username = browser.find_element_by_xpath('//*[@id="usernameTextEdit"]')
    password = browser.find_element_by_id('passwordTextEdit')
    username.clear()
    username.send_keys(config.get('SSTS', 'uname'))
    password.clear()
    password.send_keys(config.get('SSTS', 'pword'))
    browser.find_element_by_xpath('//*[@id="buttonTable"]/input').click()


def sickabs():
    global sickdate
    global enddate
    time.sleep(3)
    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
    browser.switch_to.frame('headerPlusFrame')
    browser.switch_to.frame('dataFrame')
    browser.switch_to.frame('workspaceFrame')
    browser.switch_to.frame('workspaceBodyFrame')
    browser.find_element_by_id('ListingURE_treeNode2_name').click()
    time.sleep(6)

    actionChains = ActionChains(browser)
    actionChains.double_click(browser.find_element_by_id('ListingURE_listColumn_4_0_1')).perform()
    browser.switch_to.frame('webiViewFrame')
    time.sleep(3)
    start = browser.find_element_by_xpath('//*[@id="PV1"]')
    start.clear()
    start.send_keys(sickdate.strftime('%m/%d/%Y') + " 00:00:00 AM")
    time.sleep(3)
    browser.find_element_by_xpath('//*[@id="_CWpromptstrLstElt1"]').click()
    end = browser.find_element_by_xpath('//*[@id="PV2"]')
    end.clear()
    end.send_keys(enddate.strftime('%m/%d/%Y') + " 11:59:59 PM")
    browser.find_element_by_xpath('//*[@id="_CWpromptstrLstElt2"]').click()
    browser.find_element_by_xpath('// *[ @ id = "mlst_bodyLPV3_lov"] / div / table / tbody / tr[24] / td / div').click()
    browser.find_element_by_xpath('//*[@id="theBttnIconPV3AddButton"]').click()
    browser.find_element_by_xpath('//*[@id="theBttnCenterImgpromptsOKButton"]').click()

    try:
        WebDriverWait(browser, 90).until(
            ec.presence_of_element_located((By.XPATH, '//*[@id="IconImg_iconMenu_arrow_docMenu"]')))

    except TimeoutException:
        print("Loading took too much time!")
    browser.find_element_by_xpath('//*[@id="IconImg_iconMenu_arrow_docMenu"]').click()
    hov = browser.find_element_by_id('iconMenu_menu_docMenu_span_text_saveReportComputerAs')
    ActionChains(browser).move_to_element(hov).perform()
    time.sleep(1)
    browser.find_element_by_id('saveReportComputerAs_span_text_saveReportXLS').click()

    while not os.path.exists(filename):
        time.sleep(1)

    os.rename(filename, path + "sick leave " + sickdate.strftime('%b %y') + " - " +
              enddate.strftime('%b %y') + '.xls')
    print('loop complete' + str(sickdate) + ' - ' + str(enddate))
    sickdate = sickdate - relativedelta(months=3)
    enddate = sickdate + relativedelta(months=2) + MonthEnd(1)

    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
    browser.switch_to.alert.accept()


def boxilogin():
    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/logon.jsp')
    time.sleep(2)
    browser.switch_to.frame('infoView_home')
    username = browser.find_element_by_xpath('//*[@id="usernameTextEdit"]')
    password = browser.find_element_by_id('passwordTextEdit')
    username.clear()
    username.send_keys(config.get('BOXI', 'uname'))
    password.clear()
    password.send_keys(config.get('BOXI', 'pword'))
    browser.find_element_by_xpath('//*[@id="buttonTable"]/input').click()


def boxi_employee_extracts():
    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
    browser.switch_to.frame('headerPlusFrame')
    browser.switch_to.frame('dataFrame')
    browser.switch_to.frame('workspaceFrame')
    browser.switch_to.frame('workspaceBodyFrame')
    actionChains = ActionChains(browser)
    actionChains.double_click(browser.find_element_by_id('ListingURE_listColumn_2_0_1')).perform()

    try:
        WebDriverWait(browser, 10).until(
            ec.frame_to_be_available_and_switch_to_it('webiViewFrame'))
    except TimeoutException:
        print("Loading took too much time!")
    browser.find_element_by_xpath('//*[@id="IconImg_Txt_refresh"]').click()
    try:
        WebDriverWait(browser, 10).until(
            ec.presence_of_element_located((By.XPATH, '// *[ @ id = "PV1"]')))
    except TimeoutException:
        print("Loading took too much time!")
    bdate = browser.find_element_by_xpath('// *[ @ id = "PV1"]')
    bdate.clear()
    bdate.send_keys(wstats18.strftime('%m/%d/%Y') + " 12:00:00 AM")
    browser.find_element_by_id("theBttnCenterImgpromptsOKButton").click()
    try:
        WebDriverWait(browser, 150).until(
            ec.presence_of_element_located((By.XPATH, '//*[@id="IconImg_iconMenu_arrow_docMenu"]')))
    except TimeoutException:
        print("Loading took too much time!")
    browser.find_element_by_xpath('//*[@id="IconImg_iconMenu_arrow_docMenu"]').click()
    hov = browser.find_element_by_id('iconMenu_menu_docMenu_span_text_saveReportComputerAs')
    ActionChains(browser).move_to_element(hov).perform()
    time.sleep(1)
    browser.find_element_by_id('saveReportComputerAs_span_text_saveReportXLS').click()
    print(path+'WSTATS_EMPLOYEE_EXTRACT'+'.xls')
    while not os.path.exists(path+'WSTATS_EMPLOYEE_EXTRACT.xls'):
        time.sleep(1)
    print("Extract Complete")

    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
    browser.switch_to.alert.accept()


def boxi_excess_extract():
    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
    browser.switch_to.frame('headerPlusFrame')
    browser.switch_to.frame('dataFrame')
    browser.switch_to.frame('workspaceFrame')
    browser.switch_to.frame('workspaceBodyFrame')
    actionchains = ActionChains(browser)
    actionchains.double_click(browser.find_element_by_id('ListingURE_listColumn_3_0_1')).perform()

    try:
        WebDriverWait(browser, 10).until(
            ec.frame_to_be_available_and_switch_to_it('webiViewFrame'))
    except TimeoutException:
        print("Loading took too much time!")
    browser.find_element_by_xpath('//*[@id="IconImg_Txt_refresh"]').click()
    try:
        WebDriverWait(browser, 10).until(

            ec.presence_of_element_located((By.XPATH, '// *[ @ id = "PV1"]')))
    except TimeoutException:
        print("Loading took too much time!")
    dateinput = browser.find_element_by_id('LPV1_textField')
    firstelem = browser.find_element_by_xpath('//*[@id="mlst_bodyPV1List"]/div/table/tbody/tr[1]/td/div')
    secondelem = browser.find_element_by_xpath('//*[@id="mlst_bodyPV1List"]/div/table/tbody/tr[5]/td/div')
    actionchains = ActionChains(browser)
    actionchains.key_down(Keys.SHIFT).click(firstelem).click(secondelem).key_up(Keys.SHIFT).perform()
    browser.find_element_by_xpath('//*[@id="theBttnIconPV1DelButton"]').click()

    for i in finweeks:
        dateinput.clear()
        dateinput.send_keys(i)
        browser.find_element_by_id("theBttnIconPV1AddButton").click()
    dateinput.clear()
    dateinput.send_keys(finmonth)
    browser.find_element_by_id("theBttnIconPV1AddButton").click()
    browser.find_element_by_id("theBttnCenterImgpromptsOKButton").click()
    try:
        WebDriverWait(browser, 250).until(
            ec.presence_of_element_located((By.XPATH, '//*[@id="IconImg_iconMenu_arrow_docMenu"]')))

    except TimeoutException:
        print("Loading took too much time!")
    browser.find_element_by_xpath('//*[@id="IconImg_iconMenu_arrow_docMenu"]').click()
    hov = browser.find_element_by_id('iconMenu_menu_docMenu_span_text_saveReportComputerAs')
    ActionChains(browser).move_to_element(hov).perform()
    time.sleep(1)
    browser.find_element_by_id('saveReportComputerAs_span_text_saveReportXLS').click()
    print(path + 'WSTATS_EXCESS_EXTRACT' + '.xls')
    while not os.path.exists(path + 'WSTATS_EXCESS_EXTRACT.xls'):
        time.sleep(1)
    print("Extract Complete")
    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
    browser.switch_to.alert.accept()


def boxi_bank_extract():
    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
    browser.switch_to.frame('headerPlusFrame')
    browser.switch_to.frame('dataFrame')
    browser.switch_to.frame('workspaceFrame')
    browser.switch_to.frame('workspaceBodyFrame')
    actionChains = ActionChains(browser)
    actionChains.double_click(browser.find_element_by_id('ListingURE_listColumn_1_0_1')).perform()

    try:
        WebDriverWait(browser, 10).until(
            ec.frame_to_be_available_and_switch_to_it('webiViewFrame'))

    except TimeoutException:
        print("Loading took too much time!")
    browser.find_element_by_xpath('//*[@id="IconImg_Txt_refresh"]').click()
    try:
        WebDriverWait(browser, 10).until(

            ec.presence_of_element_located((By.XPATH, '// *[ @ id = "PV1"]')))
    except TimeoutException:
        print("Loading took too much time!")
    dateinput = browser.find_element_by_id('LPV1_textField')
    firstelem = browser.find_element_by_xpath('//*[@id="mlst_bodyPV1List"]/div/table/tbody/tr[1]/td/div')
    secondelem = browser.find_element_by_xpath('//*[@id="mlst_bodyPV1List"]/div/table/tbody/tr[5]/td/div')
    actionChains = ActionChains(browser)
    actionChains.key_down(Keys.SHIFT).click(firstelem).click(secondelem).key_up(Keys.SHIFT).perform()
    browser.find_element_by_xpath('//*[@id="theBttnIconPV1DelButton"]').click()

    for i in finweeks:
        dateinput.clear()
        dateinput.send_keys(i)
        browser.find_element_by_id("theBttnIconPV1AddButton").click()
    dateinput.clear()
    dateinput.send_keys(finmonth)
    browser.find_element_by_id("theBttnIconPV1AddButton").click()
    browser.find_element_by_id("theBttnCenterImgpromptsOKButton").click()
    try:
        WebDriverWait(browser, 250).until(
            ec.presence_of_element_located((By.XPATH, '//*[@id="IconImg_iconMenu_arrow_docMenu"]')))

    except TimeoutException:
        print("Loading took too much time!")
    browser.find_element_by_xpath('//*[@id="IconImg_iconMenu_arrow_docMenu"]').click()
    hov = browser.find_element_by_id('iconMenu_menu_docMenu_span_text_saveReportComputerAs')
    ActionChains(browser).move_to_element(hov).perform()
    time.sleep(1)
    browser.find_element_by_id('saveReportComputerAs_span_text_saveReportXLS').click()
    print(path + 'WSTATS_BANK_EXTRACT' + '.xls')
    while not os.path.exists(path + 'WSTATS_BANK_EXTRACT.xls'):
        time.sleep(1)
    print("Extract Complete")
    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
    browser.switch_to.alert.accept()


def boxi_overtime_extract():
    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
    browser.switch_to.frame('headerPlusFrame')
    browser.switch_to.frame('dataFrame')
    browser.switch_to.frame('workspaceFrame')
    browser.switch_to.frame('workspaceBodyFrame')
    actionChains = ActionChains(browser)
    actionChains.double_click(browser.find_element_by_id('ListingURE_listColumn_4_0_1')).perform()

    try:
        WebDriverWait(browser, 10).until(
            ec.frame_to_be_available_and_switch_to_it('webiViewFrame'))

    except TimeoutException:
        print("Loading took too much time!")
    browser.find_element_by_xpath('//*[@id="IconImg_Txt_refresh"]').click()
    try:
        WebDriverWait(browser, 10).until(

            ec.presence_of_element_located((By.XPATH, '// *[ @ id = "PV1"]')))
    except TimeoutException:
        print("Loading took too much time!")
    dateinput = browser.find_element_by_id('LPV1_textField')
    firstelem = browser.find_element_by_xpath('//*[@id="mlst_bodyPV1List"]/div/table/tbody/tr[1]/td/div')
    secondelem = browser.find_element_by_xpath('//*[@id="mlst_bodyPV1List"]/div/table/tbody/tr[5]/td/div')
    actionChains = ActionChains(browser)
    actionChains.key_down(Keys.SHIFT).click(firstelem).click(secondelem).key_up(Keys.SHIFT).perform()
    browser.find_element_by_xpath('//*[@id="theBttnIconPV1DelButton"]').click()

    for i in finweeks:
        dateinput.clear()
        dateinput.send_keys(i)
        browser.find_element_by_id("theBttnIconPV1AddButton").click()
    dateinput.clear()
    dateinput.send_keys(finmonth)
    browser.find_element_by_id("theBttnIconPV1AddButton").click()
    browser.find_element_by_id("theBttnCenterImgpromptsOKButton").click()
    try:
        WebDriverWait(browser, 250).until(
            ec.presence_of_element_located((By.XPATH, '//*[@id="IconImg_iconMenu_arrow_docMenu"]')))

    except TimeoutException:
        print("Loading took too much time!")
    browser.find_element_by_xpath('//*[@id="IconImg_iconMenu_arrow_docMenu"]').click()
    hov = browser.find_element_by_id('iconMenu_menu_docMenu_span_text_saveReportComputerAs')
    ActionChains(browser).move_to_element(hov).perform()
    time.sleep(1)
    browser.find_element_by_id('saveReportComputerAs_span_text_saveReportXLS').click()
    print(path + 'WSTATS_OVERTIME_EXTRACT' + '.xls')
    while not os.path.exists(path + 'WSTATS_OVERTIME_EXTRACT.xls'):
        time.sleep(1)
    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
    browser.switch_to.alert.accept()
    print("Extract Complete")


def annualleave():
    global leavedate
    global leaveenddate
    time.sleep(1)
    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
    browser.switch_to.frame('headerPlusFrame')
    browser.switch_to.frame('dataFrame')
    browser.switch_to.frame('workspaceFrame')
    browser.switch_to.frame('workspaceBodyFrame')
    browser.find_element_by_id('ListingURE_treeNode2_name').click()
    time.sleep(2)
    actionChains = ActionChains(browser)
    actionChains.double_click(browser.find_element_by_id('ListingURE_listColumn_4_0_1')).perform()
    browser.switch_to.frame('webiViewFrame')
    time.sleep(2)
    start = browser.find_element_by_xpath('//*[@id="PV1"]')
    start.clear()
    start.send_keys(leavedate.strftime('%m/%d/%Y') + " 00:00:00 AM")
    time.sleep(2)
    browser.find_element_by_xpath('//*[@id="_CWpromptstrLstElt1"]').click()
    end = browser.find_element_by_xpath('//*[@id="PV2"]')
    end.clear()
    end.send_keys(leaveenddate.strftime('%m/%d/%Y') + " 11:59:59 PM")
    browser.find_element_by_xpath('//*[@id="_CWpromptstrLstElt2"]').click()
    browser.find_element_by_xpath('//*[@id="mlst_bodyLPV3_lov"]/div/table/tbody/tr[5]/td/div').click()
    browser.find_element_by_xpath('//*[@id="theBttnIconPV3AddButton"]').click()
    browser.find_element_by_xpath('//*[@id="theBttnCenterImgpromptsOKButton"]').click()

    try:
        WebDriverWait(browser, 90).until(
            ec.presence_of_element_located((By.XPATH, '//*[@id="IconImg_iconMenu_arrow_docMenu"]')))

    except TimeoutException:
        print("Loading took too much time!")
    browser.find_element_by_xpath('//*[@id="IconImg_iconMenu_arrow_docMenu"]').click()
    hov = browser.find_element_by_id('iconMenu_menu_docMenu_span_text_saveReportComputerAs')
    ActionChains(browser).move_to_element(hov).perform()
    time.sleep(1)
    browser.find_element_by_id('saveReportComputerAs_span_text_saveReportXLS').click()
    while not os.path.exists(filename):
        time.sleep(1)

    os.rename(filename, path + "annual leave " + leavedate.strftime('%b %y') + " - " + leaveenddate.strftime('%b %y')
              + '.xls')
    print('loop complete' + str(leavedate) + ' - ' + str(leaveenddate))
    leavedate = leavedate - relativedelta(months=2)
    leaveenddate = leavedate + relativedelta(months=1) + MonthEnd(1)
    print(enddate)
    print(sickdate)
    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
    browser.switch_to.alert.accept()
    print('loop complete')


def allotherabs():
    global enddate
    time.sleep(3)
    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
    browser.switch_to.frame('headerPlusFrame')
    browser.switch_to.frame('dataFrame')
    browser.switch_to.frame('workspaceFrame')
    browser.switch_to.frame('workspaceBodyFrame')
    browser.find_element_by_id('ListingURE_treeNode2_name').click()
    time.sleep(2)

    actionChains = ActionChains(browser)
    actionChains.double_click(browser.find_element_by_id('ListingURE_listColumn_4_0_1')).perform()
    browser.switch_to.frame('webiViewFrame')
    time.sleep(1)
    start = browser.find_element_by_xpath('//*[@id="PV1"]')
    start.clear()
    start.send_keys(date.strftime('%m/%d/%Y') + " 00:00:00 AM")
    time.sleep(1)
    browser.find_element_by_xpath('//*[@id="_CWpromptstrLstElt1"]').click()
    end = browser.find_element_by_xpath('//*[@id="PV2"]')
    end.clear()
    end.send_keys(enddate.strftime('%m/%d/%Y') + " 11:59:59 PM")
    browser.find_element_by_xpath('//*[@id="_CWpromptstrLstElt2"]').click()
    firstelem = browser.find_element_by_xpath('//*[@id="mlst_bodyLPV3_lov"]/div/table/tbody/tr[1]/td/div')
    secondelem = browser.find_element_by_xpath('//*[@id="mlst_bodyLPV3_lov"]/div/table/tbody/tr[33]/td/div')
    actionChains = ActionChains(browser)
    actionChains.key_down(Keys.SHIFT).click(firstelem).click(secondelem).key_up(Keys.SHIFT).perform()
    browser.find_element_by_xpath('//*[@id="theBttnIconPV3AddButton"]').click()
    firstelem = browser.find_element_by_xpath('//*[@id="mlst_bodyPV3List"]/div/table/tbody/tr[5]/td/div')
    secondelem = browser.find_element_by_xpath('//*[@id="mlst_bodyPV3List"]/div/table/tbody/tr[24]/td/div')
    actionChains = ActionChains(browser)
    actionChains.key_down(Keys.CONTROL).click(firstelem).click(secondelem).key_up(Keys.CONTROL).perform()

    browser.find_element_by_id('theBttnIconPV3DelButton').click()

    browser.find_element_by_xpath('//*[@id="theBttnCenterImgpromptsOKButton"]').click()

    try:
        WebDriverWait(browser, 90).until(
            ec.presence_of_element_located((By.XPATH, '//*[@id="IconImg_iconMenu_arrow_docMenu"]')))

    except TimeoutException:
        print("Loading took too much time!")
    browser.find_element_by_xpath('//*[@id="IconImg_iconMenu_arrow_docMenu"]').click()
    hov = browser.find_element_by_id('iconMenu_menu_docMenu_span_text_saveReportComputerAs')
    ActionChains(browser).move_to_element(hov).perform()
    time.sleep(1)
    browser.find_element_by_id('saveReportComputerAs_span_text_saveReportXLS').click()

    while not os.path.exists(filename):
        time.sleep(1)

    os.rename(filename, path + "all other leave " + date.strftime('%b %y') + '.xls')

    browser.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
    browser.switch_to.alert.accept()


boxilogin()
boxi_bank_extract()
boxi_overtime_extract()
boxi_excess_extract()
boxi_employee_extracts()
try:
    WebDriverWait(browser, 10).until(
        ec.frame_to_be_available_and_switch_to_it('headerPlusFrame'))

except TimeoutException:
    print("Loading took too much time!")

browser.find_element_by_id('btnLogoff').click()
login()
allotherabs()
for i in range(4):
    sickabs()
for i in range(2):
    annualleave()

print('Extracts Complete')
