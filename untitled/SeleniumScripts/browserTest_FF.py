
__author__='SeanPark_ViaSat'

import datetime
import time
import os
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
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

# driver = webdriver.Chrome("C:\\Selenium\\chromedriver.exe")
# driver = webdriver.Ie("C:\\Selenium\\IEDriverServer.exe")

caps = DesiredCapabilities().FIREFOX

caps["marionette"] = False

print(caps)

driver = webdriver.Firefox(capabilities=caps, executable_path="C:\\Selenium\\geckodriver.exe")

driver.get("http://www.python.org")

# driver.set_page_load_timeout(30)

# driver.get("https://ordermgmt.test.exede.net/PublicGUI-SupportGUI/v1/pages/addcustomer/serviceAvailability.xhtml")
#
# driver.get('https://spyglass01.test.wdc1.wildblue.net:8443/SpyGlass/')

installGUI = "https://igui-installationgui.test.wdc1.wildblue.net/InternalGUI-InstallationGUI/"

installGUIwithMac = installGUI + "?n=" + "AABBCCCB4149"

driver.get(installGUIwithMac)

serviceAgreementNumber = '402907978'

screenshotDirectory = './Reports/' + serviceAgreementNumber + '_' + driver.name

if not os.path.exists(screenshotDirectory):
    os.makedirs(screenshotDirectory)

driver.maximize_window()

time.sleep(3)

print("Web Browser in test : " + driver.name)

### if it's IE, it needs to bypass security warning
if driver.name == "internet explorer":
    print("driver is IE")
    continueLink = driver.find_element_by_id('overridelink')
    continueLink.click()

driver.implicitly_wait(20)

activationKeyField = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:activationKey"]'))
)

activationKeyField.send_keys(serviceAgreementNumber)

installButton = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id36"]'))
)

driver.save_screenshot(screenshotDirectory + '/1_welcomeToServiceActivation.png')

installButton.click()

time.sleep(2)

installerNumberField = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:installerId"]'))
)

installerNumberField.send_keys("99072761")

driver.save_screenshot(screenshotDirectory+'/2_customerConfirmationNewInstallation.png')

continueInstallButton = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id53"]'))
)

continueInstallButton.click()

time.sleep(5)

emailConfirmationButton = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id30"]'))
)

driver.save_screenshot(screenshotDirectory+'/3_emailConfirmationAndUpdate.png')

emailConfirmationButton.click()

time.sleep(10)

# qOIcontinueButton = WebDriverWait(driver, 60).until(
#     EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id50"]'))
# )

thankYouTag = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id40"]'))
)

print('thankYouTag Text : ' + thankYouTag.text)

qOIcontinueButton = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//input[@type="submit"]'))
)

driver.save_screenshot(screenshotDirectory+'/4_qualityOfInstall.png')

qOIcontinueButton.click()

time.sleep(6)

customerButton = WebDriverWait(driver, 60).until(
    EC.element_to_be_clickable((By.XPATH, '//input[@value="Customer"]'))
)

customerButton.click()

driver.save_screenshot(screenshotDirectory + '/5_newCustomerAccountSetup.png')

lastFourField = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:paymentAuthentication"]'))
)

lastFourField.send_keys("7777")

ccContinueButton = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//input[@value="Continue"]'))
)

ccContinueButton.click()

time.sleep(5)

# driver.switch_to_frame(1)

pdfIFrame = driver.find_element_by_xpath('//*[@id="installerForm:j_id20"]/iframe')

# print(pdfIFrame.get_attribute('src'))

driver.switch_to_default_content()

driver.switch_to_frame(pdfIFrame)

time.sleep(5)

driver.save_screenshot(screenshotDirectory + '/6_customerAgreement.png')

time.sleep(2)

getStartedButton = WebDriverWait(driver, 60).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="pnlElectronic"]/div/div[1]/button[1]/i'))
)

print("getStartedButtonAttribute : " + getStartedButton.get_attribute('class'))

getStartedButton.click()

time.sleep(3)

signField = driver.find_element_by_xpath('//*[@id="location1"]/div[2]/div[1]/input')

print(signField.get_attribute('type'))

signField.send_keys("Spider Man")

time.sleep(3)

driver.save_screenshot(screenshotDirectory + '/7_customerAgreementAfterSign.png')

finishSubmitButton = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="completePopupContainer"]/div/div[1]/button'))
)

finishSubmitButton.click()

time.sleep(2)

driver.save_screenshot(screenshotDirectory + '/8_eSignSubmitted.png')

time.sleep(2)

driver.switch_to_default_content()

continueButtonAfterSign = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id25"]'))
)

print('continueButtonAfterSign attribute : ' + continueButtonAfterSign.get_attribute('class'))

driver.save_screenshot(screenshotDirectory + '/9_eSignComplete.png')

continueButtonAfterSign.click()

time.sleep(3)

driver.save_screenshot(screenshotDirectory + '/10_activatingModem.png')

activatingModemContinueButton = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id35"]'))
)

activatingModemContinueButton.click()

time.sleep(2)

confirmationMessage = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="installerForm:j_id19"]'))
)

driver.save_screenshot(screenshotDirectory + '/11_confirmation.png')

# driver.save_screenshot('./Reports/'+'402907760'+'_provisioned.png')

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
