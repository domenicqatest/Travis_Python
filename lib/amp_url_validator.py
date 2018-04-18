#-*- coding: utf-8 -*-

from selenium import webdriver
import time
import xlrd
import xlwt
from xlutils.copy import copy


def setup_module(module):
    """The setup_module runs only one time."""
    global driver, log
    driver = webdriver.Chrome()
    driver.set_window_size(1675, 875)
    # hide window while scheduler is running
    driver.set_window_position(-2000, 0)
    driver.implicitly_wait(1)

def teardown_module(module):
    driver.quit()

class TestAmpURLs(object):

    def test_validate_urls(self):
        "This will open the AMP Validator URL and test predetermined URLs from an Excel workbook"

        # open the validator page
        driver.get("https://validator.ampproject.org/")
        time.sleep(1)

        # write excel
        #wb2 = xlrd.open_workbook('/Users/dsorace/PycharmProjects/hearst/2017/amp_validation/utils/AMP_URLs.xls')
        wb2 = xlrd.open_workbook('/Users/dsorace/PycharmProjects/hearst/2017/amp_validation/utils/AMP_URLs_STAGE.xls')

        worksheet = wb2.sheet_by_name('URLs')


        # worksheet for writing (xlwt)
        copy_wb2 = copy(wb2)

        copy_wb2.get_sheet(0).col(0).width = 25000
        copy_wb2.get_sheet(0).col(1).width = 4000

        # WRITE COLUMN NAMES
        style = xlwt.XFStyle()

        #background color
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map['gray25']
        style.pattern = pattern

        # font
        font = xlwt.Font()
        font.bold = True
        style.font = font

        write_style = copy_wb2.get_sheet(0)
        write_style.write(0, 0, "URLS", style=style)
        write_style.write(0, 1, "RESULTS", style=style)

        for row_index in range(1, worksheet.nrows):
            #site_name = worksheet.row(row_index)[1].value
            url = worksheet.row(row_index)[0].value

            # click the validator text field
            driver.find_element_by_class_name("paper-input").click()
            time.sleep(1)

            # enter the URLs one at a time
            driver.find_element_by_xpath("//input[@id='input']").send_keys(url)
            time.sleep(1)

            # click "VALIDATE" button
            driver.find_element_by_id("validateButton").click()
            time.sleep(5)

            # results
            status = driver.find_element_by_xpath("//webui-statusbar[@id='statusBar']/paper-material/div/span").text

            if status == "PASS":
                #print('\n')  # adds line break
                #print "Row Number - ", row_index, '"', site_name, '"', " validation has PASSED. - ", url

                worksheet2 = copy_wb2.get_sheet(0)

                # background color - GREEN
                style = xlwt.XFStyle()
                pattern = xlwt.Pattern()
                pattern.pattern = xlwt.Pattern.SOLID_PATTERN
                pattern.pattern_fore_colour = xlwt.Style.colour_map['light_green']
                style.pattern = pattern

                worksheet2.write(row_index, 1, status, style)
                pass
            else:
                worksheet2 = copy_wb2.get_sheet(0)

                # background color - RED
                style = xlwt.XFStyle()
                pattern = xlwt.Pattern()
                pattern.pattern = xlwt.Pattern.SOLID_PATTERN
                pattern.pattern_fore_colour = xlwt.Style.colour_map['red']
                style.pattern = pattern

                worksheet2.write(row_index, 1, status, style)

                print('\n')  # adds line break
                print "Row Number - ", row_index, " validation has FAILED. - ", url

                ### Soft Assert

            copy_wb2.save('amp_validation_results/AMP_URL_Results.xls')

            driver.find_element_by_xpath("//input[@id='input']").clear()
            time.sleep(1)

            ### Print # of Passes and Fails

    def teardown_class(self):
        driver.quit()