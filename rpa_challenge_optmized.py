# -*- coding: utf-8 -*-
"""
RPA Challenge.

Lauch RPA Challenge, download the excel file
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
import openpyxl as workbook
from decouple import config


def open_browser(_url: str) -> WebDriver:
    """
    Open new browser instance.
    """
    _browser_instance = webdriver.Firefox()
    _browser_instance.get(_url)

    return _browser_instance


def wait_for_element(_browser_instance: webdriver, _time: int, _element: str):
    """
    Wait for an element on web page to be loaded.
    """
    WebDriverWait(_browser_instance, _time).until(
        EC.presence_of_element_located((By.XPATH, _element))
    )


def download_file(_api_url: str, _file_dir: str) -> str:
    """
    Send a Get request. Return document contents in binary string and save it locally.
    """
    _response = requests.get(_api_url)

    if _response.status_code == 200:
        with open(_file_dir, 'wb') as _file:
            _file.write(_response.content)
        return str(_response.status_code)
    else:
        return str(_response.status_code) + ', ' + str(_response.text)


def read_workbook(
    _file_dir: str,
    _worksheet: str,
    _min_row: int,
    _max_row: int,
    _min_col: int,
    _max_col: int,
    _values_only: bool
    ) -> list:
    """
    Read worksheet data and return it.
    """
    if _worksheet is None:
        _worksheet = 'Sheet1'

    _wb = workbook.load_workbook(_file_dir)
    _wb_data = [
        row for row in _wb['Sheet1'].iter_rows(
            min_row = _min_row,
            max_row = _max_row,
            min_col = _min_col,
            max_col = _max_col,
            values_only = _values_only
            ) if row[0] is not None
        ]
    return _wb_data


def set_field(_browser_instance: webdriver, _data: any, _col: int):
    """
    Set field that matches the columns' numbering.
    """
    if _col == 0:
        _element = '//*[@ng-reflect-name="labelFirstName"]'
    elif _col == 1:
        _element = '//*[@ng-reflect-name="labelLastName"]'
    elif _col == 2:
        _element = '//*[@ng-reflect-name="labelCompanyName"]'
    elif _col == 3:
        _element = '//*[@ng-reflect-name="labelRole"]'
    elif _col == 4:
        _element = '//*[@ng-reflect-name="labelAddress"]'
    elif _col == 5:
        _element = '//*[@ng-reflect-name="labelEmail"]'
    elif _col == 6:
        _element = '//*[@ng-reflect-name="labelPhone"]'

    _browser_instance.execute_script(
        f'''document.evaluate(
            '{_element}',
            document,
            null,
            XPathResult.FIRST_ORDERED_NODE_TYPE,
            null
            ).singleNodeValue.value = "{_data}"'''
        )


MAIN_URL = config('RPA_URL')
API_URL = MAIN_URL + '/' + config('WEP_API')
WB_NAME = config('WB_NAME')

BTN_DOWNLOAD_EXCEL = config('BTN_DOWNLOAD_EXCEL')
BTN_START = config('BTN_START')
BTN_SUBMIT = config('BTN_SUBMIT')


browser = open_browser(MAIN_URL)
wait_for_element(browser, 10, BTN_DOWNLOAD_EXCEL)

download_file(API_URL, WB_NAME)

wbRows = read_workbook(WB_NAME, None, 2, None, None, 7, True)

browser.execute_script(
    f'''document.evaluate(
        '{BTN_START}',
        document,
        null,
        XPathResult.FIRST_ORDERED_NODE_TYPE,
        null
        ).singleNodeValue.click()'''
    )

for row in wbRows:
    for colNumber, cellData in enumerate(row):
        set_field(browser, cellData, colNumber)
    browser.execute_script(
        f'''document.evaluate(
            '{BTN_SUBMIT}',
            document,
            null,
            XPathResult.FIRST_ORDERED_NODE_TYPE,
            null
            ).singleNodeValue.click()'''
        )
