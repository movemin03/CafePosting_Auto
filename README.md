# CafePosting_Auto

[![Hits](https://hits.seeyoufarm.com/api/count/incr/badge.svg?url=https%3A%2F%2Fgithub.com%2Fmovemin03%2FCafePosting_Auto.git&count_bg=%2379C83D&title_bg=%23555555&icon=&icon_color=%23E7E7E7&title=hits&edge_flat=false)](https://hits.seeyoufarm.com)

카페 반자동 포스팅 프로그램입니다.
엑셀파일은 사이트명, 사이트주소, 사용아이디, 업로드여부, 파일명 테이블로 구성.
파일명 항목에는 카테고리+순번 조합으로 작성. 예) 기본커뮤니티 01

cafeposting_auto.py 은 네이버 전용, cafeposting_auto_daum.py 는 다음 전용입니다.
cafeposting_auto.py 에 더 중심을 두고 개발되었습니다.

cafeposting_auto_daum.py에는 현재 다음과 같은 오류 사항이 있습니다.
1. 아이디 변경 시 크롬을 제어할 수 없어지는 문제
2. 느린 속도
3. 이전 글 버튼이 제대로 눌리지 않음(변칙적 발생)

설치시 필요한 항목
pip install -r requirements.txt

requirements:
openpyxl
pyperclip
pandas
selenium
webdriver_manager


# distributer_for_씽굿.py

CafePosting_Auto 실행 전 네이버, 다음, 기타 카페를 분류해주는 프로그램입니다.

requirements:
pandas
os

# merger_for_finished_cafeposting

CafePosting_Auto 실행 완료 후 결과 엑셀 파일들을 하나로 병합시켜주는 프로그램입니다.

requirements:
glob
os
openpyxl
pandas
datetime


# foldername_extracter

포스팅 완료 후 작업일지 작성시 불필요한 작업시간을 줄이기 위하여 도입되었습니다.

requirements:
os
re
pandas
