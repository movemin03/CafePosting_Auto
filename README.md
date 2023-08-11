# CafePosting_Auto

카페 반자동 포스팅 프로그램입니다.
엑셀파일은 사이트명, 사이트주소, 사용아이디, 업로드여부, 파일명 테이블로 구성.
파일명 항목에는 카테고리+순번 조합으로 작성. 예) 기본커뮤니티1

설치시 필요한 항목
pip install -r requirements.txt

requirements:
openpyxl
pyperclip
pandas
selenium
webdriver_manager

# foldername_extracter

작업일지 작성시 불필요한 작업시간을 줄이기 위하여 도입되었습니다.
CafePosting_Auto 와는 독립적인 프로그램이지만 함께 사용하면 효율성이 극대화됩니다

requirements:
re
os
pandas
