# CafePosting_Auto

카페 반자동 포스팅 프로그램입니다.
엑셀파일은 사이트명, 사이트주소, 사용아이디, 업로드여부, 파일명 테이블로 구성.
파일명 항목에는 카테고리+순번 조합으로 작성. 예) 기본커뮤니티1

통합버전의 경우에는 다음 포스팅 시 아이디 변경이 진행되면 로그인이 제대로 작동하지 않는 오류가 발생하고 있습니다.
naver_only.py 를 이용해주시길 바랍니다. 다음 카페의 경우에는 수동 진행하실 것을 권장드립니다.

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
