#-*- coding: utf-8 -*-

from amp_urls import PROD_URLS, STAGE_URLS, SENDER_EMAIL, SENDER_PASSWORD, RECIPIENT_EMAIL
from selenium import webdriver
import time
import xlrd
import xlwt
from xlutils.copy import copy
from selenium.webdriver.common.keys import Keys
from assertlib import assertEqual
from HtmlTestRunner import HTMLTestRunner

def setup_module(module):
    """ The setup_module runs only one time.
        This test grabs URLs from an excel spreadsheet,
        runs them through the AMP Validator, then logs
        the results in a separate excel file. It then
        sends that data via email to a list of recipients.
        This is much slower than the .txt test\,
        'amp_url_validation_txt.py'."""

    global driver, log
    driver = webdriver.Chrome()
    driver.set_window_size(1675, 875)
    # hide window while scheduler is running
    driver.set_window_position(-2000, 0)
    driver.implicitly_wait(1)

def teardown_module(module):
    driver.quit()

class TestAmpURLs(object):

    def validate_urls(self):
        "This will open the AMP Validator URL and test predetermined URLs by way of 'import AMP_URLS'"
        
        # open the validator page
        driver.get("https://validator.ampproject.org/")
        time.sleep(1)

        # grab blank excel
        book = xlrd.open_workbook('/Users/dsorace/PycharmProjects/hearst/2017/amp_validation/utils/blank_excel.xls')

        # create copy for writing (xlwt) and format it
        copy_book = copy(book)

        copy_book.get_sheet(0).col(0).width = 25000
        copy_book.get_sheet(0).col(1).width = 4000

        # write column names
        style = xlwt.XFStyle()

        # background color
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map['gray25']
        style.pattern = pattern

        # font
        font = xlwt.Font()
        font.bold = True
        style.font = font

        write_style = copy_book.get_sheet(0)
        write_style.write(0, 0, "URLS", style=style)
        write_style.write(0, 1, "RESULTS", style=style)

        urls = PROD_URLS
        #urls = STAGE_URLS

        for i in xrange(271): # Number of URLs = 271
            urls_list = urls[i]

            # click the validator text field
            driver.find_element_by_class_name("paper-input").click()
            time.sleep(1)

            # enter the URLs one at a time
            driver.find_element_by_xpath("//input[@id='input']").send_keys(urls_list)
            time.sleep(1)

            # click "VALIDATE" button
            driver.find_element_by_id("validateButton").click()
            time.sleep(5)

            # results
            status = driver.find_element_by_xpath("//webui-statusbar[@id='statusBar']/paper-material/div/span").text
            worksheet = copy_book.get_sheet(0)

            if status == "PASS":

                # background color - GREEN
                style = xlwt.XFStyle()
                pattern = xlwt.Pattern()
                pattern.pattern = xlwt.Pattern.SOLID_PATTERN
                pattern.pattern_fore_colour = xlwt.Style.colour_map['light_green']
                style.pattern = pattern

                worksheet.write(i + 1, 0, urls_list)
                worksheet.write(i + 1, 1, status, style)
                pass

            else:

                # background color - RED
                style = xlwt.XFStyle()
                pattern = xlwt.Pattern()
                pattern.pattern = xlwt.Pattern.SOLID_PATTERN
                pattern.pattern_fore_colour = xlwt.Style.colour_map['red']
                style.pattern = pattern

                worksheet.write(i + 1, 0, urls_list)
                worksheet.write(i + 1, 1, status, style)

                print('\n')  # adds line break
                print "URL #", i + 1, "validation has FAILED. - ", urls_list

            copy_book.save('amp_validation_results/AMP_URL_Results.xls')

            driver.find_element_by_xpath("//input[@id='input']").clear()
            time.sleep(1)

    def test_send_email(self):
        """This will send the results via email"""
        print('\n')  # adds line break

        ### Open Outlook Mail ###
        driver.get("https://outlook.office.com/")
        time.sleep(3)
        driver.find_element_by_id("i0116").clear()
        driver.find_element_by_id("i0116").send_keys(SENDER_EMAIL)
        time.sleep(3)
        driver.find_element_by_id("idSIButton9").click()
        time.sleep(3)
        driver.find_element_by_id("passwordInput").click()
        time.sleep(3)
        driver.find_element_by_id("passwordInput").clear()
        driver.find_element_by_id("passwordInput").send_keys(SENDER_PASSWORD)
        time.sleep(3)
        driver.find_element_by_id("submitButton").click()
        time.sleep(3)
        driver.find_element_by_id("idBtn_Back").click()
        time.sleep(5)

        # click new
        driver.find_element_by_id("_ariaId_23").click()
        time.sleep(10)

        driver.find_element_by_xpath(
            "//div[@id='primaryContainer']/div[4]/div/div[1]/div/div[4]/div[1]/div/div[1]/div/div/div[1]/div/button[1]/span[1]").click()
        time.sleep(3)
        driver.find_element_by_xpath(
            "//div[@id='primaryContainer']/div[4]/div/div/div/div[4]/div[3]/div/div[5]/div/div/div[3]/div[4]/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div/div/div/span/div/form/input").clear()
        driver.find_element_by_xpath(
            "//div[@id='primaryContainer']/div[4]/div/div/div/div[4]/div[3]/div/div[5]/div/div/div[3]/div[4]/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div/div/div/span/div/form/input").send_keys(RECIPIENT_EMAIL)
        time.sleep(1)
        driver.find_element_by_xpath(
            "//div[@id='primaryContainer']/div[4]/div/div/div/div[4]/div[3]/div/div[5]/div/div/div[3]/div[4]/div/div/div[2]/div[2]/div/div/div[2]/div[6]/div[2]/div/input").click()
        time.sleep(1)
        driver.find_element_by_xpath(
            "//div[@id='primaryContainer']/div[4]/div/div/div/div[4]/div[3]/div/div[5]/div/div/div[3]/div[4]/div/div/div[2]/div[2]/div/div/div[2]/div[6]/div[2]/div/input").clear()
        driver.find_element_by_xpath(
            "//div[@id='primaryContainer']/div[4]/div/div/div/div[4]/div[3]/div/div[5]/div/div/div[3]/div[4]/div/div/div[2]/div[2]/div/div/div[2]/div[6]/div[2]/div/input").send_keys(
            "AMP VALIDATION AUTOMATION RESULTS")
        time.sleep(1)
        driver.find_element_by_xpath(
            "//div[@id='primaryContainer']/div[4]/div/div/div/div[4]/div[3]/div/div[5]/div/div/div[3]/div[4]/div/div/div[2]/div[2]/div[2]/div[3]/div/div[3]/div/div[3]/div").click()
        time.sleep(3)

    def test_send_from_excel(self):
        """This pulls the data from excel and pastes the rows one at a time into the email"""
        
        # open workbook
        book = xlrd.open_workbook(
            '/Users/dsorace/PycharmProjects/hearst/2017/amp_validation/amp_validation_results/AMP_URL_Results.xls')
        worksheet = book.sheet_by_index(0)

        pass_count = 0
        count = 271  # Number of URLs = 271

        for i in xrange(count):  # Number of URLs = 271
            rownum = (i + 1)
            is_pass = worksheet.cell((rownum), 1).value

            if is_pass == "PASS":
                pass_count += 1

        header = "PASSES - {}".format(pass_count), '\n'"FAILURES - {}"'\n'.format(count - pass_count)
        time.sleep(1)

        driver.find_element_by_xpath(
            "//div[@id='primaryContainer']/div[4]/div/div/div/div[4]/div[3]/div/div[5]/div/div/div[3]/div[4]/div/div/div[2]/div[2]/div[2]/div[3]/div/div[3]/div/div[3]/div").send_keys(header)
        time.sleep(1)
        driver.find_element_by_xpath(
            "//div[@id='primaryContainer']/div[4]/div/div/div/div[4]/div[3]/div/div[5]/div/div/div[3]/div[4]/div/div/div[2]/div[2]/div[2]/div[3]/div/div[3]/div/div[3]/div").send_keys(Keys.RETURN)
        time.sleep(1)

        # print FAILED URLs
        for i in xrange(271): # Number of URLs = 271
            rownum = (i + 1)

            url = worksheet.row(rownum)[0].value
            result = worksheet.row(rownum)[1].value

            if result == "FAIL":

                results = (url, " - ", result)

                uni =  results
                no_uni = [str(item) for item in uni]
                message = "The following URLs have failed:"
                print no_uni

                driver.find_element_by_xpath(
                    "//div[@id='primaryContainer']/div[4]/div/div/div/div[4]/div[3]/div/div[5]/div/div/div[3]/div[4]/div/div/div[2]/div[2]/div[2]/div[3]/div/div[3]/div/div[3]/div").send_keys(message)
                time.sleep(1)
                driver.find_element_by_xpath(
                    "//div[@id='primaryContainer']/div[4]/div/div/div/div[4]/div[3]/div/div[5]/div/div/div[3]/div[4]/div/div/div[2]/div[2]/div[2]/div[3]/div/div[3]/div/div[3]/div").send_keys(Keys.RETURN)
                time.sleep(1)
                driver.find_element_by_xpath(
                    "//div[@id='primaryContainer']/div[4]/div/div/div/div[4]/div[3]/div/div[5]/div/div/div[3]/div[4]/div/div/div[2]/div[2]/div[2]/div[3]/div/div[3]/div/div[3]/div").send_keys(Keys.RETURN)
                time.sleep(1)
                driver.find_element_by_xpath(
                    "//div[@id='primaryContainer']/div[4]/div/div/div/div[4]/div[3]/div/div[5]/div/div/div[3]/div[4]/div/div/div[2]/div[2]/div[2]/div[3]/div/div[3]/div/div[3]/div").send_keys(no_uni)
                driver.find_element_by_xpath(
                    "//div[@id='primaryContainer']/div[4]/div/div/div/div[4]/div[3]/div/div[5]/div/div/div[3]/div[4]/div/div/div[2]/div[2]/div[2]/div[3]/div/div[3]/div/div[3]/div").send_keys(Keys.RETURN)
                time.sleep(1)

        time.sleep(1)
        driver.find_element_by_class_name("ms-Icon--mailSend").click()
        print "Email Sent!"
        time.sleep(5)

        assertEqual(pass_count, count)
        