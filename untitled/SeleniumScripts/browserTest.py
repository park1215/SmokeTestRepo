
__author__='SeanPark_ViaSat'

import datetime
import time
import string
import random
import openpyxl
import pandas as pd
import xlsxwriter
import xlwt

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import selenium.webdriver.support.ui as ui

driver=webdriver.Chrome("C:\\Selenium\\chromedriver.exe")
# driver = webdriver.Ie("C:\Selenium\IEDriverServer.exe");

driver.set_page_load_timeout(30)

# driver.get("https://ordermgmt.test.exede.net/PublicGUI-SupportGUI/v1/pages/addcustomer/serviceAvailability.xhtml")
#
# driver.get('https://spyglass01.test.wdc1.wildblue.net:8443/SpyGlass/')

installGUI = "https://igui-installationgui.test.wdc1.wildblue.net/InternalGUI-InstallationGUI/"

installGUIwithMac = installGUI + "?n=" + "AABBCC0136C2"

driver.get(installGUIwithMac)

driver.maximize_window()

driver.implicitly_wait(20)

activationKeyField = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:activationKey"]'))
)

activationKeyField.send_keys("402907738")

installButton = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id36"]'))
)

installButton.click()

installerNumberField = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:installerId"]'))
)

installerNumberField.send_keys("99072761")

continueInstallButton = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id53"]'))
)

continueInstallButton.click()

emailConfirmationButton = WebDriverWait(driver, 180).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id30"]'))
)

emailConfirmationButton.click()

qOIcontinueButton = WebDriverWait(driver, 180).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id50"]'))
)

qOIcontinueButton.click()

customerButton = WebDriverWait(driver, 180).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id27"]'))
)

customerButton.click()

ccFourDigit = WebDriverWait(driver, 180).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:paymentAuthentication"]'))
)

ccFourDigit.send_keys("7777")

ccContinueButton = WebDriverWait(driver, 180).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id35"]'))
)

ccContinueButton.click()

# driver.switch_to_frame(1)

pdfIFrame = driver.find_element_by_xpath('//*[@id="installerForm:j_id20"]/iframe')

print(pdfIFrame.get_attribute('src'))

driver.switch_to_default_content()

driver.switch_to_frame(pdfIFrame)

time.sleep(10)

print('1')
getStartedButton = WebDriverWait(driver, 180).until(
    EC.presence_of_element_located((By.ID, 'requiredLocationCount'))
)

print('2')
print(getStartedButton.get_attribute('id'))
getStartedButton.click()
time.sleep(10)
print(3)

signField = driver.find_element_by_xpath('//*[@id="location1"]/div[2]/div[1]/input')

print(signField.get_attribute('type'))

signField.send_keys("Spider Man")

time.sleep(5)

finishSubmitButton = WebDriverWait(driver, 180).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="completePopupContainer"]/div/div[1]/button'))
)

finishSubmitButton.click()

time.sleep(5)

driver.switch_to_default_content()

continueButtonAfterSign = driver.find_element_by_xpath('//*[@id="installerForm:j_id25"]')

print(continueButtonAfterSign.get_attribute('class'))

continueButtonAfterSign.click()








#
# getStartedButton = WebDriverWait(driver, 180).until(
#     EC.presence_of_element_located((By.XPATH, '//*[@id="pnlElectronic"]/div/div[1]/button[1]'))
# )
# getStartedButton.click()
#
# signatureField = WebDriverWait(driver, 180).until(
#     EC.presence_of_element_located((By.XPATH, '//*[@id="location1"]/div[2]/div[1]/input'))
# )
# signatureField.send_keys("Spider Man")

# customerAgreementContinueButton = WebDriverWait(driver, 180).until(
#     EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id25"]'))
# )
# customerAgreementContinueButton.send_keys("Spider Man")
