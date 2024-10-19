import os
import re
import sys
import time
import traceback
import getpass
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook

# 크롤링할 게시글 개수 설정 (첫 번째 게시글 제외)
CRAWL_LIMIT = 10

# 엑셀 파일 경로 및 워크시트 이름 설정
excel_file = r'C:\Users\PHJ\output\개인정보 운영대장.xlsx'
worksheet_name = '개인정보 추출 및 이용 관리'

def extract_corporate_name(full_text):
    """
    법인명 추출: "게임사업3본부 K사업팀 / 홍길동님" 중 "게임사업3본부"만 추출
    """
    if '/' in full_text:
        return full_text.split('/')[0].strip().split()[0]
    return full_text.strip().split()[0]

def extract_file_info(file_info):
    """
    파일형식 및 파일 용량 추출
    """
    # 파일명과 용량을 '&' 또는 '('로 분리
    file_match = re.match(r'(.+?)\s*(?:&|[(])\s*([\d,\.]+\s*[KMGT]?B)', file_info, re.IGNORECASE)
    if file_match:
        filename_part = file_match.group(1).strip()
        size_part = file_match.group(2).strip()
    else:
        # '&'나 '('가 없을 경우 전체 문자열에서 용량 부분만 추출
        filename_part = file_info.strip()
        size_match = re.search(r'([\d,\.]+\s*[KMGT]?B)', filename_part, re.IGNORECASE)
        if size_match:
            size_part = size_match.group(1).strip()
            filename_part = filename_part.replace(size_part, '').strip()
        else:
            size_part = ''

    # 파일형식 결정
    if '.zip' in filename_part.lower():
        file_type = 'Zip'
    elif '.xlsx' in filename_part.lower():
        file_type = 'Excel'
    else:
        file_type = ''

    # 파일 용량 추출 (예: "24.5KB")
    size_match = re.match(r'([\d,\.]+)\s*([KMGT]?B)', size_part, re.IGNORECASE)
    if size_match:
        size_numeric = size_match.group(1).replace(',', '')
        size_unit = size_match.group(2).upper()
        file_size = f"{size_numeric} {size_unit}"
    else:
        file_size = size_part

    return file_type, file_size

def find_section_text(driver, section_titles):
    """
    특정 섹션의 제목을 기반으로 해당 섹션의 내용을 추출하는 함수
    section_titles: 섹션 제목의 리스트
    """
    try:
        # 모든 <tr> 요소를 반복하면서 섹션 찾기
        tr_elements = driver.find_elements(By.XPATH, '//table//tr')
        for tr in tr_elements:
            try:
                # 각 <tr>에서 첫 번째 <td> 요소의 텍스트 추출
                td_elements = tr.find_elements(By.TAG_NAME, 'td')
                if len(td_elements) >=2:
                    th_td = td_elements[0]
                    spans = th_td.find_elements(By.TAG_NAME, 'span')
                    header_text = ''.join([span.text.strip() for span in spans])

                    for section_title in section_titles:
                        if section_title in header_text:
                            # 해당 <tr>의 두 번째 <td> 요소의 텍스트 추출
                            value_td = td_elements[1]
                            return value_td.text.strip()
            except Exception as e:
                continue
        return None
    except Exception as e:
        print(f"find_section_text 오류: {e}")
        return None