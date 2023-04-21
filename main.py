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
import re
import pyperclip
import logging
from logging import handlers
import traceback
import sys
from PyQt6.QtCore import QSize
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)


# 로그 셋팅
log_formatter = logging.Formatter(
    ' %(asctime)s -  %(levelname)s -  %(message)s')

# handler 세팅
log_handler = handlers.TimedRotatingFileHandler(
    filename='logs/auto_text_query.log', when='midnight', interval=1, encoding='UTF8')
log_handler.setFormatter(log_formatter)
log_handler.suffix = '%Y%m%d'

logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.addHandler(log_handler)

SITE_URL = "https://apps.crossref.org/SimpleTextQuery"
GET_RESULT_TIMEOUT = 1800
RESULT_XPATH = '//*[@id="mainContent2"]/div/form/font/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[5]/td/table'

chrome_options = Options()  # 크롬 브라우저 옵션 설정
chrome_options.add_experimental_option('detach', True)


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle('Auto Text Query')

        layout = QVBoxLayout()

        src_row_layout = QHBoxLayout()
        dest_row_layout = QHBoxLayout()

        layout.addLayout(src_row_layout)
        layout.addLayout(dest_row_layout)

        src_label = QLabel('원본 폴더 지정:')
        src_btn = QPushButton('찾기')
        src_btn.clicked.connect(self.get_folder_src)
        self.src_path_label = QLabel('')
        src_row_layout.addWidget(src_label)
        src_row_layout.addWidget(src_btn)
        src_row_layout.addWidget(self.src_path_label)

        dest_label = QLabel('결과 폴더 지정:')
        dest_btn = QPushButton('찾기')
        dest_btn.clicked.connect(self.get_folder_dest)
        self.dest_path_label = QLabel('')
        dest_row_layout.addWidget(dest_label)
        dest_row_layout.addWidget(dest_btn)
        dest_row_layout.addWidget(self.dest_path_label)

        process_btn = QPushButton('시작')
        process_btn.clicked.connect(self.start_process)
        layout.addWidget(process_btn)

        container = QWidget()
        container.setLayout(layout)

        self.setStatusBar(QStatusBar(self))
        # self.setFixedSize(QSize(500, 250))
        self.setCentralWidget(container)

    def get_folder_src(self):
        initial_dir = '../lab_for_auto_text_query'
        folder_path = QFileDialog.getExistingDirectory(
            self, directory=initial_dir)
        self.src_path_label.setText(folder_path)
        print(f'path for src:', folder_path)

    def get_folder_dest(self):
        initial_dir = '../lab_for_auto_text_query'
        folder_path = QFileDialog.getExistingDirectory(
            self, directory=initial_dir)
        self.dest_path_label.setText(folder_path)
        print(f'path for dest:', folder_path)

    def start_process(self):
        print('###############################################')
        logging.info(f"{'#' * 10} 처리 시작 {'#' * 10}")
        src_path = self.src_path_label.text()
        dest_path = self.dest_path_label.text()

        regex = re.compile(r'[A-Za-z]{2}\d{8}.txt')
        file_list = os.listdir(src_path)
        logging.info(f'원본 폴더내 파일 수: {len(file_list)}')

        file_list = list(
            filter(lambda x: regex.search(x), os.listdir(src_path)))
        file_number = len(file_list)
        logging.info(f'실제 처리할 파일 수: {file_number}')

        # 원본 텍스트 파일 리스트
        for index, filename in enumerate(file_list):
            try:
                logging.info(
                    f'==========> {filename} ({index + 1} / {file_number})')

                with open(Path(src_path) / filename, encoding='UTF8') as fh:
                    content = fh.read()  # 원본 내용 가져오기
                    pyperclip.copy(content)  # 원본내용을 클릭보드에 복사

                    logging.info('원본 텍스트 복사됨')

                    browser = webdriver.Chrome(
                        'chromedriver.exe', options=chrome_options)  # 브라운저 열기
                    browser.get(SITE_URL)

                    time.sleep(0.3)

                    # TAB을 5번 누르기
                    pyautogui.write(['\t', '\t', '\t', '\t', '\t',])
                    pyautogui.hotkey('ctrl', 'v')  # 복사된 원본 텍스트를 붙여놓기

                    for i in range(2):
                        time.sleep(0.2)
                        pyautogui.write(['\t', ' '])  # 체크박스 2개 체크

                    time.sleep(0.2)
                    pyautogui.write(['\t', 'enter'])  # 엔터 누르기

                    logging.info('웹 화면에서 엔터 눌렀음')

                elem = WebDriverWait(browser, GET_RESULT_TIMEOUT).until(
                    EC.presence_of_element_located((By.XPATH, RESULT_XPATH)))  # 결과 올때까지 대기
                time.sleep(1)

                logging.info('결과 조회 완료')

                inner_html = elem.get_attribute('innerHTML')  # 결과의 HTML내용 가져오기

                soup = BeautifulSoup(inner_html, 'html.parser')
                logging.info('HTML 파싱 완료')

                result_text = ''  # 치러된 결과
                rows = soup.select('tr.resultB')
                logging.info(f'결과 항목 개수: {len(rows)}')
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
                logging.debug(f'처리된 결과: {result_text}')
                time.sleep(0.5)

                with open(Path(dest_path) / filename, 'w', encoding='UTF-8') as fh:
                    fh.write(result_text)
                    logging.info('결과 텍스트를 파일에 쓰기 완료')

                browser.quit()
                logging.info('브라우저 종료')
            except:
                logging.error(traceback.format_exc())

        logging.info(f"{'#' * 10} 처리 완료 {'#' * 10}")


app = QApplication(sys.argv)

window = MainWindow()
window.show()

app.exec()
