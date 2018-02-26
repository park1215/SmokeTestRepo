
__author__='SeanPark_ViaSat'

import datetime
import time
import string
import random
import openpyxl
import pandas as pd
import xlsxwriter
import xlwt
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import selenium.webdriver.support.ui as ui

driver = webdriver.Chrome("C:\\Selenium\\chromedriver.exe")
# driver = webdriver.Ie("C:\\Selenium\\IEDriverServer.exe")

driver.set_page_load_timeout(30)

serviceAgreementNumber = '402907958'

# driver.get("https://ordermgmt.test.exede.net/PublicGUI-SupportGUI/v1/pages/addcustomer/serviceAvailability.xhtml")
#
# driver.get('https://spyglass01.test.wdc1.wildblue.net:8443/SpyGlass/')

installGUI = "https://igui-installationgui.test.wdc1.wildblue.net/InternalGUI-InstallationGUI/"

installGUIwithMac = installGUI + "?n=" + "AABBCC9DF56C"

directory = './Reports/' + serviceAgreementNumber

if not os.path.exists(directory):
    os.makedirs(directory)

driver.save_screenshot('./Reports/' + serviceAgreementNumber + '/'+'emailConfirmationAndUpdate'+'.png')

driver.get(installGUIwithMac)

print(driver.name)