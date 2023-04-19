import re
import pyautogui
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# 노트패트 활성화
notepad = pyautogui.getWindowsWithTitle('notepad++')

SITE_URL = "https://apps.crossref.org/SimpleTextQuery"
GET_RESULT_TIMEOUT = 60
RESULT_XPATH = '//*[@id="mainContent2"]/div/form/font/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[5]/td/table'

if (len(notepad) > 0):
    print('notepad찾았음')
    notepad_window = notepad[0]
    notepad_window.activate()  # notepad 활성화
    time.sleep(0.3)
    notepad_window.maximize()  # notepad 최대화
    print(notepad_window.title)
    filename_regex = re.compile(r'PB\d{8}.txt')
    mo = filename_regex.search(notepad_window.title)
    if mo != None:
        filename = mo.group()
        print(filename)
    else:
        exit()

    pyautogui.hotkey('ctrl', 'a')  # 텍스트 전체 선택
    pyautogui.hotkey('ctrl', 'c')  # 텍스트 복사

    time.sleep(0.5)

    chr_options = Options()
    chr_options.add_experimental_option("detach", True)
    browser = webdriver.Chrome(
        './chromedriver.exe', options=chr_options)
    browser.get(SITE_URL)

    browser.maximize_window()
    time.sleep(0.3)
    pyautogui.write(['\t', '\t', '\t', '\t', '\t',])  # TAB 5번 누르기
    pyautogui.hotkey('ctrl', 'v')  # 복사된 텍스트 붙여놓기
    time.sleep(0.2)
    pyautogui.write(['\t', ' '])  # 첫번째 체크박스 체크
    time.sleep(0.2)
    pyautogui.write(['\t', ' '])  # 첫번째 체크박스 체크
    time.sleep(0.2)
    pyautogui.write(['\t', 'enter'])  # Submit버튼 클릭

    try:
        elem = WebDriverWait(browser, GET_RESULT_TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, RESULT_XPATH)))
        time.sleep(1)
        inner_html = elem.get_attribute('innerHTML')
        print(inner_html)
        soup = BeautifulSoup(inner_html, 'html.parser')
        rows = soup.select('tr.resultB')
        result_text = ''
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
        with open(f'result_files/{filename}', 'w', encoding='utf-8') as fh:
            fh.write(result_text)
    except:
        pyautogui.alert('오류 발생')
    finally:
        pyautogui.alert('처리 완료')
        browser.quit()
else:
    pyautogui.alert('Notepad++프로그램을 열어두세요!')
