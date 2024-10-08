import os
from dotenv import load_dotenv
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re

def initialize_session(login_url, username, password):
    """
    로그인 세션을 초기화하고 로그인 요청을 보낸다.
    
    Args:
        login_url (str): 로그인 URL
        username (str): 사용자 이름
        password (str): 비밀번호
    
    Returns:
        session (requests.Session): 로그인된 세션 객체
        response (requests.Response): 로그인 요청의 응답 객체
    """
    session = requests.Session()
    payload = {
        'username': username,
        'password': password
    }
    response = session.post(login_url, data=payload)
    print(f"로그인 상태: {response.status_code}")
    return session, response

def request_posts_list(session, posts_url):
    """
    게시글 목록 페이지에 GET 요청을 보내고 BeautifulSoup 객체를 반환한다.
    
    Args:
        session (requests.Session): 로그인된 세션 객체
        posts_url (str): 게시글 목록 페이지 URL
    
    Returns:
        soup (BeautifulSoup): 게시글 목록 페이지의 BeautifulSoup 객체
    """
    response = session.get(posts_url)
    print(f"게시글 목록 페이지 상태: {response.status_code}")
    soup = BeautifulSoup(response.text, 'html.parser')
    return soup

def extract_post_links(soup):
    """
    게시글 목록 페이지에서 게시글 링크를 추출한다.
    
    Args:
        soup (BeautifulSoup): 게시글 목록 페이지의 BeautifulSoup 객체
    
    Returns:
        links (list): 게시글 링크의 리스트
    """
    links = soup.select('td span[id^="Author"]')
    print(f"추출된 링크 수: {len(links)}")
    return links

def extract_url_from_onclick(onclick_value):
    """
    onclick 속성에서 URL을 추출한다.
    
    Args:
        onclick_value (str): onclick 속성의 값
    
    Returns:
        full_url (str or None): 추출된 전체 URL 또는 None
    """
    match = re.search(r"'(.*?)'", onclick_value)
    if match:
        return 'https://gw.com2us.com' + match.group(1)
    return None

def extract_data_from_post(session, post_url):
    """
    게시글 페이지에서 필요한 데이터를 추출한다.
    
    Args:
        session (requests.Session): 로그인된 세션 객체
        post_url (str): 게시글 페이지의 URL
    
    Returns:
        row_data (list or None): 추출된 데이터 리스트 또는 None
    """
    print(f"처리 중인 URL: {post_url}")
    post_response = session.get(post_url)
    post_soup = BeautifulSoup(post_response.text, 'html.parser')

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
            print(f"데이터 추출 성공: {row_data}")
            return row_data
        except Exception as e:
            print(f"데이터 추출 중 오류 발생: {str(e)}")
            return None
    else:
        print("tbody를 찾을 수 없습니다.")
        return None

def save_data_to_excel(data, output_path):
    """
    추출한 데이터를 엑셀 파일로 저장한다.
    
    Args:
        data (list): 추출된 데이터 리스트
        output_path (str): 저장할 엑셀 파일 경로
    """
    if data:
        df = pd.DataFrame(data, columns=['결제일', '년', '월', '일', '법인명', '문서번호', '제목', '신청자', '합의 담당자', '링크'])
        df.to_excel(output_path, index=False)
        print("엑셀 파일 저장 완료")
    else:
        print("추출된 데이터가 없습니다.")

def main():
    # 환경 변수 로드
    load_dotenv()
    
    # 로그인 정보 및 URL 설정
    login_url = 'https://gw.com2us.com/'
    posts_url = 'https://gw.com2us.com/portal.nsf'
    username = os.getenv('GW_USERNAME')
    password = os.getenv('GW_PASSWORD')
    output_path = os.path.join(os.getcwd(), 'output', 'output.xlsx')
    
    # 디렉토리 생성
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # 세션 초기화 및 로그인
    session, login_response = initialize_session(login_url, username, password)
    
    if login_response.status_code != 200:
        print("로그인에 실패하였습니다.")
        return
    
    # 게시글 목록 페이지 요청
    soup = request_posts_list(session, posts_url)
    
    # 게시글 링크 추출
    links = extract_post_links(soup)
    
    # 데이터 추출
    data = []
    for link in links:
        post_onclick = link.get('onclick')
        if post_onclick:
            post_url = extract_url_from_onclick(post_onclick)
            if post_url:
                row = extract_data_from_post(session, post_url)
                if row:
                    data.append(row)
            else:
                print("URL을 추출할 수 없습니다.")
        else:
            print("onclick 속성이 없습니다.")
    
    # 데이터 엑셀로 저장
    save_data_to_excel(data, output_path)

if __name__ == "__main__":
    main()
