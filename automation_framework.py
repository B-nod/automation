import time
import openpyxl
from selenium import webdriver
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import excel_opertion

test_case_location = r'C:\QA deer walk\code_for_automation\facebook_test_case.xlsx'


def read_excel():
    wb = openpyxl.load_workbook(test_case_location)
    ws = wb["Sheet1"]
    for row in ws.iter_rows(min_row=2, values_only=True):
        sn, test_summary, xpath, action, value = row[0], row[1], row[2], row[3], row[4]
        action_definition(sn, test_summary, xpath, action, value)


def action_definition(sn, test_summary, xpath, action, value):
    if action == 'open_browser':
        result, remarks = open_browser(value)
    elif action == 'open_link':
        result, remarks = open_url(value)
    elif action == 'click':
        result, remarks = click(xpath)
    elif action == 'verify_text':
        result, remarks = verify_text(xpath, value)
    elif action == 'select_dropdown':
        result, remarks = select_dropdown(xpath, value)
    elif action == 'verify_title':
        result, remarks = verify_title(value)
    elif action == 'input_text':
        result, remarks = input_text(xpath, value)
    elif action == 'close_browser':
        result, remarks = close_browser()
    elif action == 'wait':
        result, remarks = wait(value)
    else:
        result = "Not Tested"
        remarks = action+"Not Supported"
    excel_opertion.write_result(sn, test_summary, result, remarks)


def close_browser():
    try:
        driver.quit()
        result, remarks = "PASS", " "
    except Exception as ex:
        result, remarks = "FAIL", ex
    return result, remarks


def wait(value):
    try:
        time.sleep(value)
        result, remarks = "PASS", " "
    except Exception as ex:
        result, remarks = "FAIL", ex
    return result, remarks


def input_text(xpath, value):
    try:
        driver.find_element(By.XPATH, xpath).send_keys(value)
        result, remarks = "PASS", " "
    except Exception as ex:
        result, remarks = "FAIL", ex
    return result, remarks


def verify_title(value):
    try:
        actual_text = driver.title
        try:
            assert actual_text == value
            result, remarks = "Pass", ""
        except AssertionError:
            result, remarks = "Fail", "Actual value is: " + actual_text + " expected value is: " + value
    except Exception as ex:
        result, remarks = "FAIL", ex
    return result, remarks


def verify_text(xpath, value):
    try:
        actual_text = driver.find_element(By.XPATH, xpath).text
        try:
            assert actual_text == value
            result, remarks = "Pass", ""
        except AssertionError:
            result, remarks = "Fail", "Actual value is: "+actual_text+" expected value is: "+value
    except Exception as ex:
        result, remarks = "FAIL", ex
    return result, remarks


def click(xpath):
    try:
        driver.find_element(By.XPATH, xpath).click()
        result, remarks = "PASS", " "
    except Exception as ex:
        result, remarks = "FAIL", ex
    return result, remarks


def select_dropdown(xpath, value):
    try:
        Select(driver.find_element(By.XPATH, xpath)).select_by_visible_text(str(value))
        result, remarks = "PASS", " "
    except Exception as ex:
        result, remarks = "FAIL", ex
    return result, remarks


def open_url(value):
    try:
        driver.get(value)
        result, remarks = "PASS", " "
    except Exception as ex:
        result, remarks = "FAIL", ex
    return result, remarks


def open_browser(value):
    try:
        global driver
        if value == 'firefox':
            s = FirefoxService(GeckoDriverManager().install())
            driver = webdriver.Firefox(service=s)
            driver.maximize_window()
            driver.implicitly_wait(10)
            result, remarks = "PASS", " "
        elif value == 'chrome':
            s = ChromeService(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=s)
            driver.maximize_window()
            driver.implicitly_wait(10)
            result, remarks = "PASS", " "
        else:
            result, remarks = "FAIL", value+"Not supported"
    except Exception as ex:
        result, remarks = "FAIL", ex
    return result, remarks


excel_opertion.clear_result()
excel_opertion.write_header()
read_excel()
