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
    
    # 업무지원 > 개인정보 파일 전송 페이지로 이동
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
        
    # 게시글 목록 가져오기
    posts = driver.find_elements(By.XPATH, '//tr[contains(@class, "dhx_skyblue")]')
    print(f"게시글 수: {len(posts)}")