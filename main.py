import unittest
from HtmlTestRunner import HTMLTestRunner
from amp_fails_only import TestAmpURLs
from amp_fails_only import TestEmailLogin
from amp_fails_only import TestSendEmails


# get all tests from ContentCreator class
amp_validation = unittest.TestLoader().loadTestsFromTestCase(TestAmpURLs)
email_login = unittest.TestLoader().loadTestsFromTestCase(TestEmailLogin)
send_emails = unittest.TestLoader().loadTestsFromTestCase(TestSendEmails)

# create a test suite combining all tests
test_suite = unittest.TestSuite([amp_validation, email_login, send_emails])

# create output
runner = HTMLTestRunner(output='Test_Results')

# run the suite
runner.run(test_suite)