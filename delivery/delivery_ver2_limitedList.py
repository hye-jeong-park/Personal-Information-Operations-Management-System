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

def extract_corporate_name(full_text):
    """법인명 추출: "컴투스 운영지원, 이다빈" 중 "컴투스"만 추출"""
    if ',' in full_text:
        return full_text.split(',')[0].split()[0]
    return full_text.split()[0]

def extract_file_info(file_info):
    """
    파일형식 및 파일 용량 추출
    예시: "(Confidential)_20241017_103738_smon_lms_target_list.zip (221KB)"
    """
    file_match = re.match(r'\(.*?\)_(.*?)\.(zip|xlsx).*?\((\d+KB)\)', file_info)
    if file_match:
        filename = file_match.group(1) + '.' + file_match.group(2)
        file_size = file_match.group(3)
        if filename.endswith('.zip'):
            file_type = 'Zip'
        elif filename.endswith('.xlsx'):
            file_type = 'Excel'
        else:
            file_type = ''
        return file_type, file_size
    else:
        # 다른 형식의 파일명이 있을 경우
        p_tags = file_info.split('\n')
        if len(p_tags) >= 2:
            filename = p_tags[0].strip()
            file_size = p_tags[1].strip()
            if filename.endswith('.zip'):
                file_type = 'Zip'
            elif filename.endswith('.xlsx'):
                file_type = 'Excel'
            else:
                file_type = ''
            return file_type, file_size
    return '', ''

def main():
    # 웹드라이버 설정
    driver = webdriver.Chrome()
    
    try:
        # 로그인 페이지로 이동
        driver.get('https://gw.com2us.com/')
        
        # 로그인 처리
        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'Username'))
        )
        password_input = driver.find_element(By.ID, 'Password')
        
        username = input('아이디를 입력하세요: ')
        password = getpass.getpass('비밀번호를 입력하세요: ')
        
        username_input.send_keys(username)
        password_input.send_keys(password)
        
        login_button = driver.find_element(By.CLASS_NAME, 'btnLogin')
        login_button.click()
        
        # 로그인 성공 여부 확인
        time.sleep(5)
        current_url = driver.current_url
        print(f"로그인 후 현재 URL: {current_url}")
        
        if 'login' in current_url.lower():
            print("로그인에 실패하였습니다.")
            driver.quit()
            sys.exit()
        
        # 결재 > 부서함 페이지로 이동
        driver.get('https://gw.com2us.com/emate_app/00001/bbs/b2307140306.nsf/view?readform&viewname=view01')
        
        # 페이지 이동 후 현재 URL 출력
        print(f"페이지 이동 후 현재 URL: {driver.current_url}")
        
        # 페이지 로딩 대기
        try:
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, 'dhx_skyblue')))
        except Exception as e:
            print("게시글 목록을 찾을 수 없습니다.")
            print(e)
            driver.quit()
            sys.exit()

if __name__ == "__main__":
    main()