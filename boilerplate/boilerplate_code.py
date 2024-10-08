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

# 게시글 링크 추출
links = soup.select('td span[id^="Author"]')
print(f"추출된 링크 수: {len(links)}")

def extract_url_from_onclick(onclick_value):
    match = re.search(r"'(.*?)'", onclick_value)
    if match:
        return 'https://gw.com2us.com' + match.group(1)
    return None

data = []
for link in links:
    post_url = link.get('onclick')
    if post_url:
        post_url = extract_url_from_onclick(post_url)
        if post_url:
            print(f"처리 중인 URL: {post_url}")
            post_response = session.get(post_url)
            post_soup = BeautifulSoup(post_response.text, 'html.parser')

            # tbody에서 특정 값 추출
            tbody = post_soup.select_one('tbody')
            if tbody:
                try:
                    # 결제일 추출
                    payment_date = tbody.select_one('tr.date td.td_point').text.strip()
                    year, month, day = payment_date.split('-')
                    
                    # 법인명 추출
                    corporation_name = tbody.select_one('td.approval_text span#titleLabel').text.strip()
                    
                    # 문서번호 추출
                    document_number = tbody.select_one('tr.docoption td').text.strip()
                    
                    # 제목 추출
                    title = tbody.select_one('td.approval_text span#titleLabel').find_next_sibling(text=True).strip()
                    
                    # 신청자 추출
                    applicant = tbody.select_one('th[scope="row"]:contains("성명") + td span#name').text.strip()
                    
                    # 합의 담당자 추출
                    approver = tbody.select_one('tr.name td.td_point').text.strip()
                    
                    # 데이터 추가
                    row_data = [payment_date, year, month, day, corporation_name, document_number, title, applicant, approver, post_url]
                    data.append(row_data)
                    print(f"데이터 추출 성공: {row_data}")
                except Exception as e:
                    print(f"데이터 추출 중 오류 발생: {str(e)}")
            else:
                print("tbody를 찾을 수 없습니다.")
        else:
            print("URL을 추출할 수 없습니다.")