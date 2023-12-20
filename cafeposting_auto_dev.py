from urllib.parse import urljoin
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
import time
import pyperclip
import pandas as pd
import subprocess
from datetime import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import autoit
from bs4 import BeautifulSoup
from PIL import Image
import re

# 사용자가 환경에 따라 변경해야 할 값
upper_path = "" # 예)홍길동, 안에 11월 등 월별 폴더 존재해야
ver = str("2023-11-11")
auth_dic = {'id':'pw'}
chrome_ver = 119
filter_list = ['사이트명', '사이트주소', '사용아이디', '업로드여부', '파일명']
daum_id = ['exception'] # 잘 쓰지 않는 기능. 보통 다음 아이디를 여기에 넣어둠. 로그인 과정 건너뒴
except_site = ["exception"] # 사진 별도로 올릴 항목은 여기에 추가. 사진을 전부 별도로 올리려면 naver 입력

# 크롬드라이버 디버깅 모드 실행
user = os.getlogin()
subprocess.Popen(
    r'C:\Program Files\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\\Users\\' + user + r'\\AppData\\Local\\Google\\Chrome\\User Data"')
print("\n")
option = Options()
option.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driver = webdriver.Chrome(options=option)
driver.execute_script('window.open("about:blank", "_blank");')
tabs = driver.window_handles

print("\n")
print("씽굿 프로그램입니다. 네이버 전용")
print("https://github.com/movemin03/CafePosting_Auto")
print("ver:" + ver)

print("업로드 없이 카페의 회원 수 정보와 카페 랭킹 정보를 취합하는 조사 통계 모드를 사용할 수 있습니다")
print("조사통계모드로 진행할까요? 그럴거라면 1 입력, 아니면 아무거나 입력")
total_mode = str(input())

# 정보 입력
print('필요한 정보를 기입해야합니다\n')
now = time.localtime()
today_month = str(now.tm_mon).zfill(2)
today_day = str(now.tm_mday).zfill(2)
matching_folders = []

file_path = os.path.join(upper_path, today_month + "월")

if os.path.exists(file_path):
    for folder_name in os.listdir(file_path):
        if today_month + today_day in folder_name:
            folder_path = os.path.join(file_path, folder_name)
            matching_folders.append(folder_path)
    if len(matching_folders) >= 2:
        print("오늘 날짜 기준 2개 이상입니다.")
        for folder_path in matching_folders:
            print(folder_path)
            for file_name in os.listdir(folder_path):
                file_link = os.path.join(folder_path, file_name)
                if os.path.islink(file_link):
                    print(file_link)
        print("참고해서 최상위 폴더 경로를 넣어주세요")
        upper_name = input().replace('"', '')
    else:
        for upper_name in matching_folders:
            print("폴더 경로:", upper_name)
            print("가 발견되었습니다. 최상위 폴더를 변경하려면 n, 아니라면 아무거나 입력")
            a = input()
            if a == "n":
                upper_name = input().replace('"', '')
            else:
                pass
else:
    if total_mode == "1":
        pass
    else:
        print("최상위 폴더 경로를 지정해주세요")
        upper_name = input().replace('"', '')

blank_auth_dic = {}
slash_auth_dic = {}
# 슬레쉬 제거
for key, value in auth_dic.items():
    left_key = key.split('/')[0]
    slash_auth_dic[left_key] = value
auth_dic.update(slash_auth_dic)
# 공백 제거
for key, value in auth_dic.items():
    new_key = key.strip()
    blank_auth_dic[new_key] = value
auth_dic.update(blank_auth_dic)

content_path = None
title = None
post_title_path = None
img_path = None
member_txt = None
level_text = None
html_success = None

try:
    for filename in os.listdir(upper_name):
        if filename.endswith(".HTML") and filename != "mod.HTML":
            content_path = os.path.join(upper_name, filename)
            html_success = 1
except:
    if total_mode == "1":
        pass
    else:
        print("오늘 날짜가 포함된 폴더가 없거나 찾을 수 없어서 자동인식이 되지 않았습니다")
        print("최상위 폴더 경로를 지정해주세요")
        upper_name = input().replace('"', '')
        html_success = 0

def extract_html_info(content_path):
    # 파일 열기
    with open(content_path, "r", encoding="utf-8") as file:
        # HTML 파싱
        html_content = file.read()
        soup = BeautifulSoup(html_content, "html.parser")

        # 파일명에서 "_" 문자열과 "(" 문자열이 있는지 확인
        file_name = content_path.split("\\")[-1]
        if "_" in file_name and "(" in file_name:
            # "_" 문자열과 "(" 문자열 사이의 내용 추출
            start_index = file_name.index("_") + 1
            end_index = file_name.index("(")
            html_title = file_name[start_index:end_index]
        elif "(" in file_name:
            # "(" 이전까지의 내용 추출
            end_index = file_name.index("(")
            html_title = file_name[:end_index]
        else:
            # 해당 사항 없음
            html_title = ""

        # "공모전" 문자열이 있는지 확인
        if "공모전" in html_title:
            html_title = "[추천공모전] " + html_title
        else:
            html_title = "[추천대외활동] " + html_title

        html_date = None
        # "접수 기간", "접수기간", "공모기간", "공모 기간" 키워드를 포함하는 요소 탐색
        target_keywords = ["접수 기간", "접수기간", "공모기간", "공모 기간"]
        for keyword in target_keywords:
            element = soup.find(string=lambda text: text and keyword in text)
            if element:
                # 기간 추출 (예: "공모기간 : 2023년 10월 26일 ~ 11월 30일")
                date_range = element.split(":")[-1].strip()
                end_date = date_range.split("~")[-1].strip()
                html_date = "(~" + end_date.replace('년', '.').replace('(', '').replace(')', '').replace('월 ', '.').replace('일', '') +")"
                break

        return html_title + html_date

def image_compression(img_path):
    # 파일 경로에서 디렉토리와 파일명 분리
    path, file_name = os.path.split(img_path)

    # 20KB와 50KB 이미지 파일이 이미 있는지 확인
    file_20kb = os.path.join(path, "20kb_" + file_name)
    file_50kb = os.path.join(path, "50kb_" + file_name)

    if os.path.isfile(file_20kb) and os.path.isfile(file_50kb):
        print(f"20kb_display.jpg 와 50kb_display.jpg 이미지가 이미 존재합니다")
    else:
        print(f"20kb_display.jpg 와 50kb_display.jpg 파일이 없어서 변환합니다")
        # 이미지를 불러와서 압축
        img = Image.open(img_path).convert('RGB')

        quality = 90  # 초기 quality 값을 설정
        img.save(file_50kb, "JPEG", quality=quality)

        new_file_size = os.path.getsize(file_50kb) / 1024  # KB 단위로 변환

        # 파일 크기가 50KB 이하가 될 때까지 quality를 줄여가며 압축
        while new_file_size > 50:
            quality -= 10  # quality 값을 줄임
            img.save(file_50kb, "JPEG", quality=quality)
            new_file_size = os.path.getsize(file_50kb) / 1024  # KB 단위로 변환

        # 20KB 이하로 압축하는 과정
        quality = 50  # 초기 quality 값을 설정
        img.save(file_20kb, "JPEG", quality=quality)

        new_file_size = os.path.getsize(file_20kb) / 1024  # KB 단위로 변환

        # 파일 크기가 20KB 이하가 될 때까지 quality를 줄여가며 압축
        while new_file_size > 20:
            quality -= 5  # quality 값을 줄임
            img.save(file_20kb, "JPEG", quality=quality)
            new_file_size = os.path.getsize(file_20kb) / 1024  # KB 단위로 변환

if total_mode == "1":
    pass
else:
    if not html_success == 1:
        for filename in os.listdir(upper_name):
            if filename.endswith(".HTML"):
                content_path = os.path.join(upper_name, filename)

    for filename in os.listdir(upper_name):
        if filename == "제목.txt":
            post_title_path = os.path.join(upper_name, filename)

    for filename2 in os.listdir(upper_name):
        if filename2 == "display.jpg" or filename2 == "display.jpeg" or filename2 == "display.png":
            img_path = os.path.join(upper_name, filename2)
            break

    if content_path:
        print(f"HTML 파일 경로: {content_path}")

    else:
        print("HTML 파일을 찾을 수 없습니다. .HTML 인지 확인해주십시오. .html 처럼 소문자인 경우에도 인식하지 못합니다")
        print("아래에 수동으로 입력해주세요:")
        content_path = input().replace('"', '')


    if post_title_path:
        print(f"제목.txt 파일 경로: {post_title_path}")
    else:
        print("제목.txt 파일을 찾을 수 없습니다.")
        try:
            html_info = extract_html_info(content_path)
            print("html 파일을 참고하여 추천해드리겠습니다")
            print("파일별로 접수기간을 제시하는 형식이 모두 달라서 자동으로 인식되는 내용이 잘못되었을 수 있습니다")
            print(html_info)
            print("위 내용이 맞는지 확인하고 맞으면 엔터 아니라면 n 입력")
            a = input()
            if a == "n":
                print("제목.txt 파일에 들어갈 내용을 입력해주시면 제목.txt로 저장됩니다")
                html_info = input()
            else:
                pass
            with open(upper_name + "\\제목.txt", "w", encoding="utf-8") as file:
                file.write(html_info)
            print("제목.txt 를 생성합니다")
            for filename in os.listdir(upper_name):
                if filename == "제목.txt":
                    post_title_path = os.path.join(upper_name, filename)
        except:
            pass


    if img_path:
        print(f"이미지 파일 경로: {img_path}")
        image_compression(img_path)
    else:
        print("업로드 할 이미지 파일을 찾을 수 없습니다")

        try:
            print("이미지를 자동 내려받기를 시도합니다")
            with open(content_path, 'r', encoding='utf-8') as file:
            # 파일 내용 읽기
                file_content = file.read()

            # img 태그 찾기
                img_tags = file_content.split('<img')

            # 이미지 다운로드 및 저장
                for img_tag in img_tags:
                    if 'src="' in img_tag:
                    # 이미지 링크 추출
                        img_src_start = img_tag.index('src="') + len('src="')
                        img_src_end = img_tag.index('"', img_src_start)
                        img_src = img_tag[img_src_start:img_src_end]

                    # 이미지 링크를 절대 경로로 변환
                        img_url = urljoin(content_path, img_src)
                        img_url = img_url.replace("amp;", "")

                    # 이미지 다운로드
                        response = requests.get(img_url)

                    # 이미지 파일명 추출
                        img_filename = os.path.basename(img_url)

                    # 이미지 저장 파일명
                        save_filename = 'display.jpg'
                        save_path = os.path.join(upper_name, save_filename)

                    # 이미지 저장
                        with open(save_path, 'wb') as img_file:
                            img_file.write(response.content)

                        print('이미지를 성공적으로 저장했습니다.')
                        img_path = save_path
                        print(img_path)
                        print('이미지 위치를 지정했습니다.')
                        image_compression(img_path)
        except:
            print("업로드 할 이미지 파일을 찾을 수 없습니다. jpg png jpeg 파일을 적용 가능합니다")
            print("아래에 수동으로 입력해주세요:")
            img_path = input().replace('"', '')
            image_compression(img_path)

    try:
        with open(post_title_path, encoding='utf-8') as file:
            title = file.readline().strip()
        print("업로드할 게시물 이름:", title)
    except:
        print("게시물의 제목을 적어주십시오:")
        title = input()

try:
    for filename in os.listdir(upper_name):
        if filename == "naver_cafe.xlsx":
            upload_path = os.path.join(upper_name, filename)
except:
    pass

try:
    print(f"\n참고 엑셀 파일 경로: {upload_path}")
    print("참고할 엑셀 파일 경로 변경을 원하시면 n 을 아니면 아무거나 입력")
    a = input()
    if a == "n":
        print('참고할 엑셀 파일 위치를 알려주세요:')
        upload_path = input().replace('"', '')
    else:
        pass
except:
    print('참고할 엑셀 파일 위치를 알려주세요:')
    upload_path = input().replace('"', '')
    print(f"참고 엑셀 파일 경로: {upload_path}")



# 데이터 전처리0
print('\n 입력하신 엑셀파일을 읽어오고 있습니다')
excel = pd.read_excel(upload_path, names=['사이트명', '사이트주소', '사용아이디', '업로드여부', '파일명'])
pd.set_option('mode.chained_assignment', None)
input_id_list = list(excel['사용아이디'].drop_duplicates())
try:
    input_id_list = [item.strip() for item in input_id_list]
    input_id_list = [string.split("/")[0].strip() for string in input_id_list]
except:
    print("100 퍼센트의 확률로 제공하신 엑셀파일의 아이디 칸이 비어 있습니다. 프로그램을 재실행하시길 권장합니다")

n_error_list = []
wait = WebDriverWait(driver, 5)
error_posting_url = 0
start_num = 0

def remove_img_tags(upper_name, content_path_mod):
    with open(upper_name, 'r', encoding='utf-8') as file:
        html_content = file.read()

    soup = BeautifulSoup(html_content, 'html.parser')

    # <img> 태그 제거
    for img_tag in soup.find_all('img'):
        img_tag.decompose()

    modified_html = str(soup)

    with open(content_path_mod, 'w', encoding='utf-8') as file:
        file.write(modified_html)

except_site_num = 0
def check_url(n_url, except_site):
    for site in except_site:
        if site in n_url:
            print("이 사이트는 html 붙여넣기 외에 별도로 사진을 업로드 하도록 지정하셨습니다")
            wait.until(EC.presence_of_element_located((By.XPATH, '//p[contains(@class,"se-text-paragraph se-text-paragraph-align-left")]')))
            driver.find_elements(By.XPATH, '//p[contains(@class,"se-text-paragraph se-text-paragraph-align-left")]')[0].click()
            action = ActionChains(driver)
            action.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
            action.send_keys(Keys.BACKSPACE)
            # 사진 업로드 선택
            time.sleep(1)
            driver.find_element(By.XPATH, '//button[contains(@class,"se-image-toolbar-button se-document-toolbar-basic-button se-text-icon-toolbar-button")]').click()
            time.sleep(2)
            handle = "[CLASS:#32770; TITLE:열기]"
            img_path.replace(";", ":")
            print("img_path를 붙여넣기 합니다: " + img_path)
            autoit.control_send(handle, "Edit1", img_path)
            time.sleep(1)
            autoit.control_click(handle, "Button1")
            time.sleep(1)
            print("사진이 업로드 되었습니다")
            global content_path_mod
            content_path_mod = upper_name + "\\mod.HTML"
            if "mod.HTML" in os.listdir(upper_name):
                print("mod.HTML 파일이 이미 있습니다")
            else:
                print("기존 html의 태그를 일부 수정하여 mod.HTML 파일을 만들고 있습니다")
                print("시간이 조금 소요될 수 있습니다")
                remove_img_tags(content_path, content_path_mod)
                print("파일 생성이 완료되었습니다")
            global except_site_num
            except_site_num = 1
    pass
def content_html():
    global except_site_num
    if except_site_num == 1:
        driver.get(content_path_mod)
        except_site_num = 0
    else:
        driver.get(content_path)
    action = ActionChains(driver)
    action.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
    action.key_down(Keys.CONTROL).send_keys('c').key_up(Keys.CONTROL).perform()
def login():
    global tabs
    tabs = driver.window_handles
    print(str(tabs))
    try:
        if len(tabs) >= 2:
            driver.switch_to.window(tabs[1])
        else:
            print("탭이 부족해서 새 탭을 엽니다")
            driver.switch_to.window(tabs[0])
            driver.execute_script('window.open("about:blank", "_blank");')
            tabs = driver.window_handles
            driver.switch_to.window(tabs[1])
            time.sleep(2)
    except:
        print("탭 갯수: " + str(len(tabs)))
        print("탭 관리에 문제가 있습니다")
        print("탭을 수동으로 열어주십시오")
        a = input()

    login_url = "https://nid.naver.com/nidlogin.login"
    driver.get(login_url)
    print("로그인 링크 접속 완료: " + login_url)

    pyperclip.copy(auth)
    driver.find_element(By.ID, 'id').click()
    action = ActionChains(driver)
    action.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
    action.send_keys(Keys.BACKSPACE).perform()
    driver.find_element(By.ID, 'id').send_keys(Keys.CONTROL + 'v')

    try:
        password = auth_dic[auth]
    except:
        print("auth_dic을 가져오는 것에 오류가 있어서 패스워드를 수동입력해야 합니다:")
        password = input()
    pyperclip.copy(password)
    driver.find_element(By.ID, 'pw').click()
    action.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
    action.send_keys(Keys.BACKSPACE).perform()
    driver.find_element(By.ID, 'pw').send_keys(Keys.CONTROL + 'v')

    time.sleep(1)
    login_btn = driver.find_element(By.ID, 'log.login')
    login_btn.click()
    print("로그인: 로그인 작업 진행 완료\n")
    a = input("2차 인증 여부 및 아이디가 " + auth + "가 맞는지 확인해주시고 아무거나 입력 후 엔터")

def total():
    global status_stop
    status_stop = 0

    global tabs
    tabs = driver.window_handles
    print(str(tabs))
    try:
        if len(tabs) >= 2:
            driver.switch_to.window(tabs[1])
        else:
            print("탭이 부족해서 새 탭을 엽니다")
            driver.switch_to.window(tabs[0])
            driver.execute_script('window.open("about:blank", "_blank");')
            tabs = driver.window_handles
            driver.switch_to.window(tabs[1])
            time.sleep(2)
    except:
        print("탭 갯수: " + str(len(tabs)))
        print("탭 관리에 문제가 있습니다")
        print("탭을 수동으로 열어주십시오")
        a = input()

    if naver_url.count("/") > 3:
        split_text = naver_url.split("/")
        split_text = split_text[:4]  # 처음 4개의 요소만 유지
        n_url = "/".join(split_text)
    else:
        n_url = naver_url
    driver.get(n_url)
    print("링크 접속 완료: " + n_url)
    global error_myactivity
    error_myactivity = 0
    global error_myactivity_detail
    error_myactivity_detail = ""

    # 멤버수, 카페 등급 조회
    print("회원수 및 사이트 등급 정보를 불러옵니다")
    wait.until(EC.presence_of_element_located((By.XPATH, '//li[@class="level-info border-sub"]/*[1]/*[3]')))
    member = driver.find_element(By.XPATH, '//li[@class="level-info border-sub"]/*[1]/*[3]')
    global member_txt
    member_txt = member.text
    print(str(member_txt))

    level = driver.find_element(By.XPATH,'//li[@class="mem-cnt-info"]/a/*[2]')
    global level_text
    level_text = level.text.replace("비공개", "").replace("공개", "")
    print(str(level_text))
def posting():
    global status_stop
    status_stop = 0

    global tabs
    tabs = driver.window_handles
    print(str(tabs))
    try:
        if len(tabs) >= 2:
            driver.switch_to.window(tabs[1])
        else:
            print("탭이 부족해서 새 탭을 엽니다")
            driver.switch_to.window(tabs[0])
            driver.execute_script('window.open("about:blank", "_blank");')
            tabs = driver.window_handles
            driver.switch_to.window(tabs[1])
            time.sleep(2)
    except:
        print("탭 갯수: " + str(len(tabs)))
        print("탭 관리에 문제가 있습니다")
        print("탭을 수동으로 열어주십시오")
        a = input()

    if naver_url.count("/") > 3:
        split_text = naver_url.split("/")
        split_text = split_text[:4]  # 처음 4개의 요소만 유지
        n_url = "/".join(split_text)
    else:
        n_url = naver_url
    driver.get(n_url)
    print("링크 접속 완료: " + n_url)
    global error_myactivity
    error_myactivity = 0
    global error_myactivity_detail
    error_myactivity_detail = ""

    # 포스팅
    try:
        driver.find_element(By.CLASS_NAME, 'cafe-info-action').find_element(By.XPATH,
                                                                            '//button[contains(text(),"나의활동")]').click()
        wait.until(EC.presence_of_element_located((By.XPATH, '//a[contains(text(),"내가 쓴 게시글")]')))
        driver.find_element(By.XPATH, '//a[contains(text(),"내가 쓴 게시글")]').click()
    except:
#네이버 카페 업데이트로 하위 경로 탐색기능이 작동되지 않음
        try:
            driver.find_element(By.CLASS_NAME, 'cafe-info-action').find_element(By.XPATH,
                                                                                '//a[contains(text(),"나의활동")]').click()
            wait.until(EC.presence_of_element_located((By.XPATH, '//a[contains(text(),"내가 쓴 글 보기")]')))
            driver.find_element(By.XPATH, '//a[contains(text(),"내가 쓴 글 보기")]').click()
        except:
            try:
                driver.find_element(By.CLASS_NAME, 'cafe-info-action').find_element(By.XPATH,
                                                                                    '//a[contains(text(),"나의활동")]').click()
                wait.until(EC.presence_of_element_located((By.XPATH, '//a[contains(text(),"내가 쓴 글 보기")]')))
                driver.find_element(By.XPATH, '//a[contains(text(),"내가 쓴 글 보기")]').click()
            except:
                print(
                    '오류: 가입되지 않은 카페 or 강퇴 or (낮은 확률로 활동정지)에 의한 오류입니다. 아이디가 ' + auth + '가 맞는지 확인하고 아니라면 재로그인해주세요. 이후 아무키나 눌러주십시오')
                a = input()
                try:
                    driver.find_element(By.CLASS_NAME, 'cafe-info-action').find_element(By.XPATH,
                                                                                        '//button[contains(text(),"나의활동")]').click()
                    wait.until(EC.presence_of_element_located((By.XPATH, '//a[contains(text(),"내가 쓴 게시글")]')))
                    driver.find_element(By.XPATH, '//a[contains(text(),"내가 쓴 게시글")]').click()
                    error_myactivity = 0
                except:
                    print('오류: 가입되지 않은 카페 or 강퇴 or (낮은 확률로 활동정지)에 의한 오류')
                    error_myactivity = 1
                    print("어떤 오류인지 알려줄 수 있나요? 1. 미가입 2. 잘못된 경로 3. 강퇴 4. 활동정지 5. 금칙어 6. 기타")
                    error_myactivity_detail = input()

                    error_messages = {
                    '1': "미가입 오류",
                    '2': "잘못된 경로",
                    '3': "강퇴",
                    '4': "활동정지",
                    '5': "금칙어 사용 오류",
                    '6': "탭 오류",
                    '7': "기타 오류",}

                    error_myactivity_detail = (error_messages[error_myactivity_detail])

    # frame으로 변환해야 게시글 확인이 가능하다#

    if error_myactivity == 0:
        time.sleep(1)
        former_post = 1
        try:
            driver.switch_to.frame("cafe_main")
            mcn = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/table/tbody/tr[1]/td[1]/div[3]/div/a')
            mcn_lnk = mcn.get_attribute('href')
            driver.get(mcn_lnk)

        except:
            if error_myactivity == 0:
                print('이전 게시물이 없어서 프로그램이 참조할 것이 없습니다.')
                former_post = 0
            else:
                pass

        driver.switch_to.default_content()
        wait.until(EC.presence_of_element_located((By.NAME, "cafe_main")))
        driver.switch_to.frame("cafe_main")
        # 이전 포스트가 있는 경우 이전 포스트 내 글쓰기 누르고, 없는 경우정지상태 새 글 작성

        if former_post == 1:
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/div/div[3]/div[1]/a[1]')))
                mcn = driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[3]/div[1]/a[1]')
                mcn_lnk = mcn.get_attribute('href')
                driver.get(mcn_lnk)
            except:
                print("높은 확률로 정지 상태입니다. 간혹 등급 부족인 경우에도 이 문구가 뜰 수 있습니다")
                status_stop = 1
        else:
            wait.until(EC.presence_of_element_located((By.XPATH, '//span[contains(@class,"BaseButton__txt")]')))
            mcn = driver.find_element(By.XPATH, '//span[contains(@class,"BaseButton__txt")]').find_element(By.XPATH,
                                                                                                           './..')
            mcn_lnk = mcn.get_attribute('href')
            driver.get(mcn_lnk)

        try:
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[2]/div[1]/div[1]/div/div[2]/div/textarea')))
            driver.find_element(By.XPATH,'//*[@id="app"]/div/div/section/div/div[2]/div[1]/div[1]/div/div[2]/div/textarea').click()

            action = ActionChains(driver)
            action.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
            action.send_keys(title).perform()
            print("글쓰기 1/4: 제목 입력 완료")

            #붙여넣기 이외에 사진을 별도로 업로드해야하는 경우인지 아닌지 판단하고, 추가 작업을 진행한다
            check_url(n_url, except_site)
            print("글쓰기 2/3: 사진 별도 업로드 필요 여부 확인 완료")

            driver.switch_to.window(tabs[0])
            content_html()
            driver.switch_to.window(tabs[1])
            wait.until(EC.presence_of_element_located(
                (By.XPATH, '//p[contains(@class,"se-text-paragraph se-text-paragraph-align-left")]')))
            driver.find_elements(By.XPATH, '//p[contains(@class,"se-text-paragraph se-text-paragraph-align-left")]')[
                0].click()
            action.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
            print("글쓰기 3/4: html 코드 입력 완료")

            if former_post == 0:
                try:
                    try:
                        wait.until(EC.presence_of_element_located(
                            (By.XPATH, '//button[contains(text(),"게시판을 선택해 주세요")]')))
                        driver.find_element(By.XPATH, '//button[contains(text(),"게시판을 선택해 주세요")]').click()
                        elements = driver.find_element(By.CLASS_NAME, 'option_list')
                    except:
                        pass
                    try:
                        print(elements.text)
                        print("카테고리는 위와 같습니다")
                    except:
                        print("[자유] 키워드가 포함된 카테고리가 없는 것으로 보입니다")
                    print('카테고리를 수동 변경해야 합니다. 카테고리를 설정해주세요. 수정 후 아무키나 눌러주세요')
                except:
                    pass
                a = input()
            else:
                time.sleep(3)

            driver.find_element(By.XPATH, '//span[contains(@class,"BaseButton__txt")]').click()
            print("글쓰기 4/4: 업로드 완료")

            try:
                alert = driver.switch_to.alert  # alert 창에 대한 참조를 가져옵니다.
                alert_text = alert.text  # alert 창에 표시된 텍스트를 가져옵니다.
                print(str(alert_text))
                alert.accept()
                time.sleep(4)
                driver.find_element(By.XPATH, '//button[contains(text(),"등록")]').click()
                print("글쓰기 3/3: 업로드 완료: 연속된 글쓰기로 감지되었음")
            except:
                pass

            time.sleep(1)
            try:
                try:
                    driver.switch_to.default_content()
                    wait.until(EC.presence_of_element_located((By.NAME, "cafe_main")))
                    driver.switch_to.frame("cafe_main")
                    global wait_fail
                    wait_fail = 0
                except:
                    print("대기시간이 소용이 없군!")
                    print('포스팅 완료 과정에서 오류가 있었습니다. 대부분 금칙어 오류. 종종 글쓰기 활동 중지 혹은 연속된 게시물인 경우 오류가 발생할 수 있습니다')
                    error_myactivity_detail = "대부분 금칙어 오류. 종종 글쓰기 활동 중지 혹은 연속된 게시물인 경우 오류가 발생할 수 있습니다"
                    wait_fail = 1
                    pass
                global posting_url_n
                posting_url_n = "NaN"
                if wait_fail == 1:
                    posting_url_n = "오류가 있습니다: " + naver_url
                else:
                    pass
                posting_url_n = str(driver.current_url)
                if "articles/write" in posting_url_n:
                    posting_url_n = f"잘못된 링크가 들어갔으므로 수동 작업 필요:write: {posting_url_n}"
                    global error_posting_url
                    error_posting_url = 1
                else:
                    if posting_url_n == "NaN" or posting_url_n == "":
                        posting_url_n = f"잘못된 링크가 들어갔으므로 수동  작업 필요:NaN: + {naver_url}"
                        error_posting_url = 1
                    else:
                        pass
            except:
                posting_url_n = "(금칙어 처리) 혹은 연속된 게시물 또는 창이 꺼짐. 본래 url: " + naver_url
                error_posting_url = 1
            print("포스팅된 게시물 링크 posting_url: " + posting_url_n)
            time.sleep(2)
        except:
            posting_url_n = "높은 확률로 정지 상태(간혹 등급부족 오류)혹은 프로그램 멈춤. 본래 url: " + naver_url
            print("높은 확률로 정지 상태입니다. 둥급 부족인 상태에도 해당 문구가 뜰 수 있습니다")
            print("프로그램 멈춤 시에도 해당 오류가 발생합니다")
            status_stop = 1
    else:
        posting_url_n = "오류: 가입되지 않은 카페 or 강퇴 or (낮은 확률로 활동정지)에 의한. 본래 url: " + naver_url

# 실행되는 라인
print('본 프로그램에 등록되어 있는 id 는 다음과 같습니다:' + str(list(auth_dic.keys())))
if set(input_id_list) - set(list(auth_dic.keys())) == set():
    print('모두 자동 업로드 가능합니다')
    pre_list = str(input_id_list)
else:
    print('요청하신 id는 다음과 같습니다: ' + str(input_id_list))
    same = set(list(auth_dic.keys())) & set(input_id_list)
    if same == set():
        print('현재 자동 업로드를 진행할 수 있는 아이디가 없습니다')
        exit()
    else:
        print('요청하신 내용의 일부인 ' + str(same) + '를 자동 업로드할 수 있습니다')
        pre_list = str(same)
    print('auth_dic 에 다음 값을 추가해야 합니다 : ' + str(set(input_id_list) - set(list(auth_dic.keys()))))
    print('추가 없이 진행 시, 추가하지 않은 id 는 업로드하지 않습니다\n')
    a = input
# 불필요한 기호 생략
pre_list = pre_list.translate(str.maketrans("", "", "{}',"))

List = [x for x in pre_list.split()]
for x in List:
    auth = x
    auth = auth.replace("[", "")
    auth = auth.replace("]", "")
    auth = auth.replace("[", "")
    print("사용할 아이디: " + auth)
    execute_time = datetime.today().strftime("%Y%m%d_%H%M")

    # 데이터 전처리
    excel_1 = excel[excel['사용아이디'] == auth]
    print(excel_1)

    url_list = list(excel_1['사이트주소'])

    naver_list = [x for x in url_list if "cafe.naver.com" in x]
    len_naver = len(naver_list)

    if len_naver > 0:
        try:
            if auth in daum_id or total_mode == "1":
                print("조사 모드이거나 다음아이디로 판별되어 로그인 과정을 건너 뜁니다")
                excel_1["카페등급"] = ""
                excel_1["회원수"] = ""
            else:
                login()
        except:
            print('오류: 이미 로그인이 진행 or 자동로그인 기능 겹침 or id/pw 오류. 수동 진행 후 아무 키 입력. 수동 진행 후 아무 키 입력')
            a = input()
    else:
        print('업로드할 naver 링크가 없습니다')
    print("\n" + auth + " 아이디로 네이버 카페 " + str(len_naver) + "개를 진행하겠습니다")
    i = 0
    while i < len_naver:
        try:
            naver_url = naver_list[i]
            if auth in daum_id or total_mode == "1":
                total()
                try:
                    if member_txt is not None:
                        excel_1.iloc[i, 5] = member_txt
                        excel_1.iloc[i, 6] = level_text
                    else:
                        print("member_txt 값이 없습니다")
                except Exception as e:
                    print(e)
                    print("회원수 및 등급 정보 입력을 실패했습니다")
            else:
                posting()
            if error_posting_url == 1:
                excel_1.iloc[i, 3] = "X"
                excel_1.iloc[i, 1] = posting_url_n + " : " +error_myactivity_detail
            else:
                if wait_fail == 1:
                    excel_1.iloc[i, 3] = "X"
                    excel_1.iloc[i, 1] = posting_url_n + " : " +error_myactivity_detail
                else:
                    if error_myactivity == 1:
                        excel_1.iloc[i, 3] = "X"
                        excel_1.iloc[i, 1] = posting_url_n + " : " +error_myactivity_detail
                    else:
                        if status_stop == 1:
                            excel_1.iloc[i, 3] = "X"
                            excel_1.iloc[i, 1] = posting_url_n + " : " +error_myactivity_detail
                        else:
                            excel_1.iloc[i, 3] = "O"
                            excel_1.iloc[i, 1] = posting_url_n
                            if "1" in excel_1.iloc[i, 4] and not re.search(r"[023456789]|11|111", excel_1.iloc[i, 4]):
                                print("파일명에 1 이 포함되어 스크린샷을 진행합니다")
                                screenshot_path = upper_name +  "\\" + str(excel_1.iloc[i, 0]) + "_" + str(excel_1.iloc[i, 4]) + ".png"
                                original_size = driver.get_window_size()
                                print("저장된 값인 900,1080의 창크기, 화면비율 73퍼센트로 자동 조절합니다")
                                driver.set_window_size(900, 1080)
                                driver.execute_script("document.body.style.zoom = '73%'")
                                try:
                                    actions = ActionChains(driver)
                                    actions.send_keys(Keys.PAGE_UP)
                                    actions.perform()
                                except Exception as e:
                                    print("오류 메시지:", str(e))

                                print("크롬 창크기를 조정해주세요 크기대로 캡쳐됩니다. 하고 엔터")
                                a = input()
                                driver.save_screenshot(screenshot_path)
                                print(screenshot_path + "에 스크린샷이 저장되었습니다")
                                driver.set_window_size(original_size['width'], original_size['height'])
                                driver.execute_script("document.body.style.zoom = '100%'")
                            else:
                                pass
            posting_url_n = "NaN"
            error_posting_url = 0
        except:
            if i >= len_naver:
                pass
            else:
                n_error_list.append(naver_list[i])
                if status_stop == 1:
                    excel_1.iloc[i, 3] = "X-활동정지"
                else:
                    excel_1.iloc[i, 3] = "X"
        i = i + 1
    naver_list = []
    len_naver = 0

    # 결과 폴더 경로 지정
    result_dir = 'C:\\Users\\' + user + '\\Desktop\\결과'

    # 결과 폴더가 없으면 생성
    if not os.path.exists(result_dir):
        os.makedirs(result_dir)
    excel_1.to_excel(result_dir + "\\" + auth + "_" + execute_time + '_작업완료naver.xlsx')

print('작업완료된 내역을 엑셀 파일로 저장하였습니다.')
