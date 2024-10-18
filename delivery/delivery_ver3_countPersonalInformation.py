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