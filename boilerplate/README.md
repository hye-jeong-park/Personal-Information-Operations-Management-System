# 제목 없음

# GW Portal Data Extraction

이 프로젝트는 Com2uS GW 포털에서 특정 게시글 데이터를 추출하여 엑셀 파일로 저장하는 Python 스크립트입니다. 로그인, 페이지 요청, 데이터 파싱 및 엑셀 파일 생성의 단계를 포함하고 있습니다.

## 기능

- **로그인 정보 관리**: `.env` 파일을 통해 사용자 이름과 비밀번호를 관리하여 로그인 세션을 유지합니다.
- **게시글 목록 페이지 크롤링**: BeautifulSoup을 사용하여 게시글 링크를 추출합니다.
- **게시글 상세 정보 파싱**: 각 게시글에서 결제일, 법인명, 문서번호, 제목, 신청자, 합의 담당자 등의 정보를 추출합니다.
- **엑셀 파일 저장**: 추출한 데이터를 엑셀 파일로 저장합니다.

## 요구 사항

- Python 3.x
- 필수 라이브러리: `requests`, `beautifulsoup4`, `pandas`, `python-dotenv`

필수 라이브러리는 아래 명령어로 설치할 수 있습니다:

```bash
bash
코드 복사
pip install requests beautifulsoup4 pandas python-dotenv
```

## 사용법

1. .env 파일 설정
    
    프로젝트 폴더에 `.env` 파일을 생성하고 다음과 같이 로그인 정보를 입력합니다.
    
    ```makefile
    makefile
    코드 복사
    GW_USERNAME=your_username
    GW_PASSWORD=your_password
    ```
    
2. 스크립트 실행
    
    스크립트를 실행하여 데이터를 추출합니다:
    
    ```bash
    bash
    코드 복사
    python script_name.py
    ```
    
3. 결과 파일
    
    스크립트 실행 후 `output.xlsx` 파일이 지정된 경로(`C:\Users\PHJ\output\output.xlsx`)에 저장됩니다.
    

## 코드 설명

- **로그인 세션 유지**: `requests.Session()`을 사용하여 세션을 유지하고 로그인 후 데이터 요청을 수행합니다.
- **데이터 추출**: BeautifulSoup을 사용하여 게시글 상세 페이지의 `tbody` 영역에서 특정 데이터를 추출합니다.
- **엑셀 파일 저장**: 추출한 데이터를 `pandas` DataFrame으로 변환하여 엑셀 파일로 저장합니다.

## 예외 처리

데이터 추출 과정에서 발생할 수 있는 오류는 `try-except` 블록으로 처리됩니다. 예를 들어, 특정 데이터를 찾을 수 없거나 예상과 다른 형식일 경우 오류 메시지가 출력됩니다.

## 주의 사항

- 이 코드는 GW 포털의 특정 페이지 구조에 의존하므로, 페이지 구조가 변경되면 코드 수정을 해야 할 수 있습니다.
- `.env` 파일에는 민감한 정보가 포함되므로 파일을 안전하게 관리하고 공유하지 않도록 주의하세요.