# -*- coding: utf-8 -*-
"""
RPA Challenge.

Lauch RPA Challenge .
"""

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

browser = webdriver.Firefox()
browser.get('https://rpachallenge.azurewebsites.net/')

btnDownloadExcel = WebDriverWait(browser, 10).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a'))
)

browserAction = ActionChains(browser)
browserAction.move_to_element(btnDownloadExcel).click().perform()
