import pyautogui
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from pathlib import Path
import os
import pyperclip

LAB_PATH = Path(
    r'C:\Users\M2S\Desktop\Workspace\PythonWorkspace\lab_for_auto_text_query')
print(Path.cwd())
SRC_PATH = LAB_PATH / 'src'
DEST_PATH = LAB_PATH / 'dest'

SITE_URL = "https://apps.crossref.org/SimpleTextQuery"
GET_RESULT_TIMEOUT = 1800
RESULT_XPATH = '//*[@id="mainContent2"]/div/form/font/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[5]/td/table'

chrome_options = Options()  # 크롬 브라우저 옵션 설정
chrome_options.add_experimental_option('detach', True)

# 원본 텍스트 파일 리스트
for filename in os.listdir(SRC_PATH):
    print(filename)
    with open(SRC_PATH / filename, encoding='UTF8') as fh:
        content = fh.read()  # 원본 내용 가져오기
        pyperclip.copy(content)  # 원본내용을 클릭보드에 복사

        browser = webdriver.Chrome(
            'chromedriver.exe', options=chrome_options)  # 브라운저 열기
        browser.get(SITE_URL)

        time.sleep(0.3)

        pyautogui.write(['\t', '\t', '\t', '\t', '\t',])  # TAB을 5번 누르기
        pyautogui.hotkey('ctrl', 'v')  # 복사된 원본 텍스트를 붙여놓기

        for i in range(2):
            time.sleep(0.2)
            pyautogui.write(['\t', ' '])  # 체크박스 2개 체크

        time.sleep(0.2)
        pyautogui.write(['\t', 'enter'])  # 엔터 누르기

    elem = WebDriverWait(browser, GET_RESULT_TIMEOUT).until(
        EC.presence_of_element_located((By.XPATH, RESULT_XPATH)))  # 결과 올때까지 대기
    time.sleep(1)
    inner_html = elem.get_attribute('innerHTML')  # 결과의 HTML내용 가져오기

    soup = BeautifulSoup(inner_html, 'html.parser')

    result_text = ''  # 치러된 결과
    rows = soup.select('tr.resultB')
    print(f'lenghth of result: {len(rows)}')
    for row in rows:
        td = row.td
        count = len(td.contents)
        if count > 0:
            item = td.contents[0].string.text.strip().replace(
                '\n      ', '')
            result_text += item + '\n'
        if count > 2:
            link = td.contents[2].string.text
            result_text += link + '\n'
        if count > 4:
            bottom_id = td.contents[4].string.text
            result_text += bottom_id + '\n'
        result_text += '\n'
    time.sleep(0.5)
    with open(DEST_PATH / filename, 'w', encoding='UTF-8') as fh:
        fh.write(result_text)

    browser.quit()
