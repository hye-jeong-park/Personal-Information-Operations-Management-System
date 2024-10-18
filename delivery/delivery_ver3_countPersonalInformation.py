from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import traceback
import getpass
from openpyxl import load_workbook
import sys
import re

# 크롤링할 게시글 개수 설정 (첫 번째 게시글 제외)
CRAWL_LIMIT = 10 

# 엑셀 파일 경로 및 워크시트 이름 설정
excel_file = r'C:\Users\PHJ\output\개인정보 운영대장.xlsx'
worksheet_name = '개인정보 추출 및 이용 관리'

def extract_corporate_name(full_text):
    """
    법인명 추출: "컴투스 운영지원, 홍길동" 중 "컴투스"만 추출
    """
    if ',' in full_text:
        return full_text.split(',')[0].split()[0]
    return full_text.split()[0]

def extract_file_info(file_info):
    """
    파일형식 및 파일 용량 추출
    """
    # 파일명과 용량을 '&' 또는 ','로 분리
    if '&' in file_info:
        parts = file_info.split('&')
    elif ',' in file_info:
        parts = file_info.split(',')
    else:
        parts = [file_info]

    if len(parts) >= 2:
        filename_part = parts[0].strip()
        size_part = parts[1].strip()
    else:
        filename_part = parts[0].strip()
        size_part = ''

    # 파일형식 결정
    if '.zip' in filename_part.lower():
        file_type = 'Zip'
    elif '.xlsx' in filename_part.lower():
        file_type = 'Excel'
    else:
        file_type = ''

    # 파일 용량 추출 (예: "61,104KB" 또는 "61,104 KB")
    size_match = re.search(r'(\d{1,3}(?:,\d{3})*)\s*(KB|MB)', size_part, re.IGNORECASE)
    if size_match:
        file_size = f"{size_match.group(1).replace(',', '')}{size_match.group(2).upper()}"
    else:
        file_size = ''

    return file_type, file_size

def find_section_text(driver, section_title):
    """
    특정 섹션의 제목을 기반으로 해당 섹션의 내용을 추출하는 함수
    """
    # 모든 <tr> 요소를 반복하면서 섹션 찾기
    tr_elements = driver.find_elements(By.XPATH, '//table//tr')
    for tr in tr_elements:
        tds = tr.find_elements(By.TAG_NAME, 'td')
        if len(tds) < 2:
            continue
        try:
            # 첫 번째 <td>의 첫 번째 <span> 텍스트 확인
            header_span = tds[0].find_element(By.TAG_NAME, 'span')
            header_text = header_span.text.strip()
            if section_title == header_text:
                return tds[1].text.strip()
        except:
            continue
    return None

def main():
    # 웹드라이버 설정
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)

    try:
        # 로그인 페이지로 이동
        driver.get('https://gw.com2us.com/')
        
        # 로그인 처리
        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'Username'))
        )
        password_input = driver.find_element(By.ID, 'Password')
        
        # 사용자로부터 아이디와 비밀번호 입력받기
        username = input('아이디를 입력하세요: ')
        password = getpass.getpass('비밀번호를 입력하세요: ')
        
        # 아이디와 비밀번호 입력
        username_input.send_keys(username)
        password_input.send_keys(password)
        
        # 로그인 버튼 클릭
        login_button = driver.find_element(By.CLASS_NAME, 'btnLogin')
        login_button.click()
        
        # 로그인 성공 여부 확인
        WebDriverWait(driver, 10).until(
            EC.url_changes('https://gw.com2us.com/')
        )
        current_url = driver.current_url
        print(f"로그인 후 현재 URL: {current_url}")
        
        if 'login' in current_url.lower():
            print("로그인에 실패하였습니다.")
            driver.quit()
            sys.exit()

if __name__ == "__main__":
    main()