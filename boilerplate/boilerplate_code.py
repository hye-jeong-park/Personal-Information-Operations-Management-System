import os
from dotenv import load_dotenv
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re

# .env 파일 로드
load_dotenv()

# 로그인 세션 유지
session = requests.Session()
login_url = 'https://gw.com2us.com/'
payload = {
    'username': os.getenv('GW_USERNAME'),
    'password': os.getenv('GW_PASSWORD')
}
response = session.post(login_url, data=payload)
print(f"로그인 상태: {response.status_code}")

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
                    payment_date_full = tbody.select_one('td.td_point').text.strip()
                    payment_date = payment_date_full.split(' ')[0]  # "2024-09-05" 추출
                    year, month, day = payment_date.split('-')
                    
                    # 법인명 추출
                    corporation_name = tbody.select_one('td.approval_text span#titleLabel').text.strip()
                    
                    # 문서번호 추출
                    document_number = tbody.select_one('tr.docoption td').text.strip()
                    
                    # 제목 추출
                    title = tbody.select_one('td.approval_text').text.split(corporation_name)[-1].strip()
                    
                    # 신청자 추출
                    applicant = tbody.select_one('span#name').text.strip()
                    
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

# 데이터프레임으로 변환 후 엑셀로 저장
if data:
    df = pd.DataFrame(data, columns=['결제일', '년', '월', '일', '법인명', '문서번호', '제목', '신청자', '합의 담당자', '링크'])
    df.to_excel(r'C:\Users\PHJ\output\output.xlsx', index=False)
    print("엑셀 파일 저장 완료")
else:
    print("추출된 데이터가 없습니다.")