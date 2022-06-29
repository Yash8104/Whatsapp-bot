import os
from time import sleep
from selenium import webdriver
import openpyxl as xl
import pyautogui
import config

wb = xl.load_workbook(config.excel_file_name)
sheet = wb[config.sheet_name]
col = config.col
startrow = config.startrow
endrow = config.endrow
number = []


def gettingnumber():
    global startrow
    global number
    global endrow

    while startrow <= endrow:
        cell = sheet[col + str(startrow)]
        try:
            value = int(cell.value)
            number.append(value)
            startrow = startrow + 1
        except:
            startrow += 1
            gettingnumber()


gettingnumber()
driver = webdriver.Chrome()


def start(target):
    driver.get(f"https://web.whatsapp.com/send?phone=+91{target}")
    sleep(15)

    if config.sending_pdf:
        driver.find_element_by_css_selector('[title="Attach"]') \
            .click()
        sleep(4)
        driver.find_element_by_css_selector('[data-testid="attach-document"]') \
            .click()
        sleep(2)

        pyautogui.write(os.getcwd() + f"\\{config.pdf_file_name}")
        sleep(2)
        pyautogui.press('enter')

        sleep(2)

        driver.find_element_by_css_selector('[data-testid="send"]') \
            .click()
        sleep(10)
    else:

        driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]') \
            .send_keys(config.message)
        sleep(3)
        driver.find_element_by_css_selector('[data-testid="send"]') \
            .click()

        sleep(10)


for i in range(0, len(number)):
    start(target=number[i])

driver.quit()
