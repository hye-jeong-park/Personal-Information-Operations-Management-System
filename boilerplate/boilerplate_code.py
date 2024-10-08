import requests
from bs4 import BeautifulSoup
import pandas as pd
import re

# 로그인 세션 유지
session = requests.Session()
login_url = 'https://gw.com2us.com/'
payload = {
    'username': 'happyloopy',
    'password': '1234'
}
response = session.post(login_url, data=payload)
print(f"로그인 상태: {response.status_code}")

# 게시글 목록 페이지 요청
response = session.get('https://gw.com2us.com/portal.nsf')
soup = BeautifulSoup(response.text, 'html.parser')
print(f"게시글 목록 페이지 상태: {response.status_code}")

# 게시글 목록 페이지 요청
response = session.get('https://gw.com2us.com/portal.nsf')
soup = BeautifulSoup(response.text, 'html.parser')
print(f"게시글 목록 페이지 상태: {response.status_code}")