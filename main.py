import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
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
from PyQt6.QtCore import pyqtSignal, QThread
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QWidget,
)

basedir = os.path.dirname(__file__)

# 로그 셋팅
log_formatter = logging.Formatter(
    ' %(asctime)s -  %(levelname)s -  %(message)s')

log_path = Path(basedir) / 'logs'
if not os.path.exists(log_path):
    os.mkdir(log_path)

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
FILE_NAME_REG = r'.*\.txt'

chrome_options = Options()  # 크롬 브라우저 옵션 설정
chrome_options.add_experimental_option('detach', True)

style_sheet = """
    QProgressBar{
        background-color: #C0C6CA;
        color: #FFFFFF;
        border: 1px solid grey;
        padding: 3px;
        height: 15px;
        text-align: center;
    }

    QProgressBar::chunk{
        background: #538DB8;
        width: 5px;
        margin: 0.5px
    }
"""


class Worker(QThread):
    update_progress_bar_signal = pyqtSignal(int)
    update_text_edit_signal = pyqtSignal(str)
    clear_text_edit_signal = pyqtSignal()

    def __init__(self, src_dir, dest_dir):
        super().__init__()
        self.src_dir = src_dir
        self.dest_dir = dest_dir

    def stop_running(self):
        """Terminate the thread."""
        self.terminate()
        self.wait()

        self.update_progress_bar_signal.emit(0)
        self.clear_text_edit_signal.emit()

    def run(self):
        logging.info(f"{'#' * 10} 처리 시작 {'#' * 10}")

        regex = re.compile(FILE_NAME_REG)
        file_list = os.listdir(self.src_dir)
        logging.info(f'원본 폴더내 파일 수: {len(file_list)}')

        file_list = list(
            filter(lambda x: regex.search(x), os.listdir(self.src_dir)))
        file_number = len(file_list)
        logging.info(f'실제 처리할 파일 수: {file_number}')

        # 원본 텍스트 파일 리스트
        for index, filename in enumerate(file_list):
            try:
                logging.info(
                    f'==========> {filename} ({index + 1} / {file_number})')

                with open(Path(self.src_dir) / filename, encoding='UTF8') as fh:
                    src_text = fh.read()  # 원본 내용 가져오기
                    fh.seek(0)
                    src_line_num = len(fh.readlines())  # 원본 항목 수
                    logging.info(f'원본 항목 수: {src_line_num}')
                    pyperclip.copy(src_text)  # 원본내용을 클릭보드에 복사

                logging.info('원본 텍스트 복사됨')

                browser = webdriver.Chrome(
                    'chromedriver.exe', options=chrome_options)  # 브라운저 열기
                browser.get(SITE_URL)

                textarea = browser.find_element(
                    By.CSS_SELECTOR, '#freetext')

                textarea.click()

                action = ActionChains(browser)
                action.key_down(Keys.CONTROL).send_keys('V').key_up(Keys.CONTROL) \
                    .send_keys(Keys.TAB).send_keys(Keys.SPACE) \
                    .send_keys(Keys.TAB).send_keys(Keys.SPACE) \
                    .send_keys(Keys.TAB).send_keys(Keys.ENTER).perform()

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

                item_count_is_same = src_line_num == len(rows)

                if item_count_is_same:  # 원본 항목수 와 결과 항목수가 일치하는 경우
                    logging.info('원본 항목수와 결과 항목수 일치')
                    with open(Path(self.dest_dir) / filename, 'w', encoding='UTF-8') as fh:
                        fh.write(result_text)
                        logging.info('결과 텍스트를 파일에 쓰기 완료')
                else:  # 불일치하는 경우: 결과 폴더에 원본 텍스트를 쓰고 하위에 에러폴더를 생성해서 결과를 에러폴더에 저장
                    logging.info('원본 항목수와 결과 항목수 불일치')
                    with open(Path(self.dest_dir) / filename, 'w', encoding='UTF-8') as fh:
                        fh.write(src_text)
                        logging.info('원본 텍스트를 파일에 쓰기 완료')
                        os.makedirs(Path(self.dest_dir) /
                                    'error', exist_ok=True)  # 에러 폴더 생성
                        with open(Path(self.dest_dir) / 'error' / filename, 'w', encoding='UTF-8') as err_fh:
                            err_fh.write(result_text)
                            logging.info('결과 텍스트를 에러 폴더에 쓰기 완료')

                browser.quit()
                logging.info('브라우저 종료')
                self.update_progress_bar_signal.emit(index + 1)
                log_content = f'[INFO] ==> {filename} 처리완료 ({index + 1} / {file_number}) [{src_line_num} == {len(rows)}]' if item_count_is_same else f'[WARNING] ==> {filename} 처리완료 ({index + 1} / {file_number}) [{src_line_num} != {len(rows)}]'
                self.update_text_edit_signal.emit(log_content)
            except:
                logging.error(traceback.format_exc())
                self.update_progress_bar_signal.emit(index + 1)
                self.update_text_edit_signal.emit(
                    f'[ERROR] ==> {filename} 치리오류 ({index + 1} / {file_number})')

        # 프로그래스 바의 값을 0으로 리셋
        self.update_progress_bar_signal.emit(0)


class MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.initializeUI()

    def initializeUI(self):
        """Set up the application's GUI."""
        self.setWindowTitle('Auto Text Query')
        self.setGeometry(1300, 100, 600, 400)
        self.setFixedSize(700, 600)

        self.src_dir = ''
        self.dest_dir = ''

        self.setUpMainWindow()
        self.show()

    def setUpMainWindow(self):
        """Create and arrange widgets in the main window."""
        dir_label = QLabel(
            """<p>원본폴더 / 결과폴더를 선택하세요:</p>""")
        self.src_dir_edit = QLineEdit()
        src_dir_button = QPushButton("원본 폴더")
        src_dir_button.setToolTip("원본 폴더 선택")
        src_dir_button.clicked.connect(self.choose_directory)

        self.dest_dir_edit = QLineEdit()
        dest_dir_button = QPushButton("결과 폴더")
        dest_dir_button.setToolTip("결과  폴더 선택")
        dest_dir_button.clicked.connect(self.choose_directory)

        start_button = QPushButton('시작')
        start_button.setToolTip('처리 시작')
        start_button.clicked.connect(self.start_process)

        self.display_log_tedit = QTextEdit()
        self.display_log_tedit.setReadOnly(True)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)

        self.stop_button = QPushButton('중지')
        self.stop_button.setEnabled(False)

        # Create layout and arrange widgets
        grid = QGridLayout()

        grid.addWidget(dir_label, 0, 0)

        grid.addWidget(self.src_dir_edit, 1, 0, 1, 2)
        grid.addWidget(src_dir_button, 1, 2)

        grid.addWidget(self.dest_dir_edit, 2, 0, 1, 2)
        grid.addWidget(dest_dir_button, 2, 2)

        grid.addWidget(start_button, 3, 0, 1, 2)

        grid.addWidget(self.display_log_tedit, 4, 0, 1, 3)

        grid.addWidget(self.progress_bar, 5, 0, 1, 2)
        grid.addWidget(self.stop_button, 5, 2)

        self.setLayout(grid)

    def choose_directory(self):
        btn_text = self.sender().text()

        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.FileMode.Directory)
        selected_dir = file_dialog.getExistingDirectory(
            self, '폴더 선택', '', QFileDialog.Option.ShowDirsOnly)

        if selected_dir and '원본' in btn_text:
            self.src_dir_edit.setText(selected_dir)
            print(selected_dir)

            # 프로그래스 바의 최대값 설정
            num_of_files = len(list(filter(lambda x: re.compile(
                FILE_NAME_REG).search(x), os.listdir(selected_dir))))
            if num_of_files > 0:
                self.progress_bar.setRange(0, num_of_files)
                self.src_dir = selected_dir
            else:
                self.src_dir = ''
        elif selected_dir and '결과' in btn_text:
            self.dest_dir_edit.setText(selected_dir)
            self.dest_dir = selected_dir

    def start_process(self):
        """Create instance of worker thread to handle process."""
        if self.src_dir != '' and self.dest_dir != '':
            self.worker = Worker(self.src_dir, self.dest_dir)
            self.display_log_tedit.clear()

            self.stop_button.setEnabled(True)
            self.stop_button.repaint()
            self.stop_button.clicked.connect(self.worker.stop_running)

            self.worker.update_progress_bar_signal.connect(
                self.update_progress_bar)
            self.worker.update_text_edit_signal.connect(self.update_text_edit)
            self.worker.finished.connect(self.process_finished)
            self.worker.start()
        else:
            QMessageBox.warning(self, '알림', '원본/결과 폴더를 모두 선택하세요!',
                                QMessageBox.StandardButton.Ok)

    def update_progress_bar(self, value):
        self.progress_bar.setValue(value)

    def update_text_edit(self, log_message):
        self.display_log_tedit.append(log_message)

    def process_finished(self):
        self.display_log_tedit.append('### 처리 완료 ###')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet(style_sheet)
    app.setWindowIcon(QIcon(os.path.join(basedir, 'icon.png')))
    window = MainWindow()
    sys.exit(app.exec())
