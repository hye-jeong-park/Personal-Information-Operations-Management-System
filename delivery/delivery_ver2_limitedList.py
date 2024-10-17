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
        
        # 게시글 목록 가져오기
        posts = driver.find_elements(By.XPATH, '//tr[contains(@class, "dhx_skyblue")]')
        total_posts = len(posts)
        print(f"총 게시글 수: {total_posts}")
        
        if total_posts <= 1:
            print("처리할 게시글이 없습니다. (첫 번째 게시글만 존재)")
            driver.quit()
            sys.exit()
        
        # 크롤링할 게시글 개수 설정 (첫 번째 게시글 제외)
        limit = min(CRAWL_LIMIT, total_posts - 1)
        print(f"크롤링할 게시글 개수: {limit}")
                
        data_list = []
        
        for i in range(1, limit + 1):  # 첫 번째 게시글은 인덱스 0이므로 1부터 시작
            # 게시글 목록을 다시 가져옵니다. (동적 페이지일 경우 필요)
            posts = driver.find_elements(By.CSS_SELECTOR, 'tr[class*="dhx_skyblue"]')
            if i >= len(posts):
                print(f"게시글 {i+1}은 존재하지 않습니다. 종료합니다.")
                break
            post = posts[i]
        
            try:
                # 해당 행의 모든 td 요소를 가져옵니다.
                tds = post.find_elements(By.TAG_NAME, 'td')
        
                # 등록일 추출 (5번째 td, 0-based index)
                등록일_td = tds[4]
                등록일_text = 등록일_td.get_attribute('title').strip() if 등록일_td.get_attribute('title') else 등록일_td.text.strip()
        
                # 작성자 추출 (3번째 td)
                작성자_td = tds[2]
                작성자 = 작성자_td.find_element(By.TAG_NAME, 'span').text.strip()
        
            except Exception as e:
                print(f"목록에서 데이터 추출 중 오류 발생 (게시글 {i+1}): {e}")
                등록일_text = 작성자 = ''
                continue  # 오류 발생 시 다음 게시글로 이동
        
            # 요소가 화면에 보이도록 스크롤합니다.
            driver.execute_script("arguments[0].scrollIntoView();", post)
        
            # 클릭 가능할 때까지 대기합니다.
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(post))
            
            # 게시글 클릭하여 팝업 열기
            post.click()
        
            # 새로운 창으로 전환
            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
            driver.switch_to.window(driver.window_handles[-1])
        
            # 페이지 로딩 대기
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'AppLineArea')))
        
            try:
                # 상세 페이지에서 제목 확인
                h2_element = driver.find_element(By.CSS_SELECTOR, '#AppLineArea h2')
                h2_text = h2_element.text.strip()
        
                # 제목이 '개인정보 추출 신청서'가 아닌 경우 건너뜀
                if '개인정보 추출 신청서' not in h2_text:
                    print(f"게시글 {i+1}: 제목이 '개인정보 추출 신청서'가 아닙니다. 건너뜁니다.")
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                    continue  # 다음 게시글로 이동
        
                # 현재 창 제목 출력
                print(f"게시글 {i+1}: 현재 창 제목: {driver.title}")
        

if __name__ == "__main__":
    main()