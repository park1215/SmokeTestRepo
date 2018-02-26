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
import unittest
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import selenium.webdriver.support.ui as ui
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import HtmlTestRunner
direct = os.getcwd()

class MyTestSuite(unittest.TestCase):

    def setUp(self):
        # self.driver = webdriver.Chrome("C:\\Selenium\\chromedriver.exe")
        # driver = webdriver.Ie("C:\Selenium\IEDriverServer.exe");

        print("self in setup: " + str(self))

    # def test_search_in_python_org(self):
    #     driver = self.driver
    #
    #     driver.implicitly_wait(50)
    #
    #     driver.set_page_load_timeout(30)
    #
    #     driver.get("http://www.python.org")
    #
    #     self.assertIn("Python", driver.title)
    #
    #     print("self in test_search_in_python_org: "+str(self))
    #
    #     driver.maximize_window()
    #
    #     driver.implicitly_wait(20)
    #
    #     assert "Python" in driver.title
    #     elem = driver.find_element_by_name("q")
    #     elem.clear()
    #     elem.send_keys("pycon")
    #     elem.send_keys(Keys.RETURN)
    #     assert "No results found." not in driver.page_source
    #     # print(driver.page_source)

    # def test_search_in_python_org2(self):
    #     driver = self.driver
    #
    #     driver.implicitly_wait(50)
    #
    #     driver.set_page_load_timeout(30)
    #
    #     driver.get("http://www.python.org")
    #
    #     self.assertIn("Python", driver.title)
    #
    #     print("self in test_search_in_python_org2: "+str(self))
    #
    #     driver.maximize_window()
    #
    #     driver.implicitly_wait(20)
    #
    #     assert "Python" in driver.title
    #     elem = driver.find_element_by_name("q")
    #     elem.clear()
    #     elem.send_keys("Python")
    #     elem.send_keys(Keys.RETURN)
    #     assert "No results found." not in driver.page_source
    #     # print(driver.page_source)

    def test_upper(self):
        self.assertEqual('foo'.upper(), 'FOO')

    def test_isupper(self):
        self.assertTrue('FOO'.isupper())
        self.assertFalse('Foo'.isupper())

    def test_split(self):
        s = 'hello world'
        self.assertEqual(s.split(), ['hello', 'world'])
        # check that s.split fails when the separator is not a string
        with self.assertRaises(TypeError):
            s.split(2)

    # def test_error(self):
    #     """ This test should be marked as error one. """
    #     raise ValueError

    def test_fail(self):
        """ This test should fail. """
        self.assertEqual(2, 2)

    @unittest.skip("This is a skipped test.")
    def test_skip(self):
        """ This test should be skipped. """
        pass

    def tearDown(self):
        print("self in tearDown: " + str(self))
        # self.driver.close()

if __name__ == "__main__":
    HtmlTestRunner.main(testRunner=HtmlTestRunner.HTMLTestRunner(output='C:\\Users\\spark\\PycharmProjects\\untitled\\SeleniumScripts\\Reports\\TestReport.html'))
    # HtmlTestRunner.main()