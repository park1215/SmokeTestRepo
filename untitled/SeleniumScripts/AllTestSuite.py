__author__ = 'SeanPark_Viasat'

import unittest
import SeleniumScripts
import os
from HtmlTestRunner import HTMLTestRunner
from SeleniumScripts import TestCases
direct = os.getcwd()

class AllTestSuite(unittest.TestCase):
    def test_Issue(self):
        smoke_test = unittest.TestSuite()
        smoke_test.addTests([
            unittest.defaultTestLoader.loadTestsFromTestCase(TestCases.MyWikiTestCase),
            unittest.defaultTestLoader.loadTestsFromTestCase(TestCases.MyGoogleTestCase),
        ])

        outfile = open(direct + "\SmokeTest.html", "w")

        runner1 = HTMLTestRunner.HTMLTestRunner(
            stream=outfile,
            title='Test Report',
            description='Smoke Tests'
        )

        runner1.run(smoke_test)

if __name__ == '__main__':
    unittest.main()