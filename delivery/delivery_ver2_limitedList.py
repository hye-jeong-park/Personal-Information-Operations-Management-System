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