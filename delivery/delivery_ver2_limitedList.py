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