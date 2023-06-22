from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
import time
import pyperclip
import pandas as pd
import subprocess

subprocess.Popen(r'C:\Program Files\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\\Users\\movem\\AppData\\Local\\Google\\Chrome\\User Data"')
option = Options()
option.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
chrome_ver = 114
driver = webdriver.Chrome(options=option)
driver.execute_script('window.open("about:blank", "_blank");')
tabs = driver.window_handles

ver = str("2023-06-22")

# 안내
print("\n")
print("씽굿 프로그램입니다.")
print("https://github.com/movemin03/CafePosting_Auto")
print("ver:" + ver)

# 정보 입력
print('필요한 정보를 기입해야합니다\n')
auth_dic = {'id': 'pw'}
print("게시물의 제목을 적어주십시오:")
title = input()
print("복사할 html 코드 위치를 입력해주세요:")
content_path = input()
try:
    content_path = content_path.replace('"', '')
except:
    pass

# 데이터 전처리0
print('참고할 엑셀 파일 위치를 알려주세요:')
upload_path = input()
try:
    upload_path = upload_path.replace('"', '')
except:
    pass
print('\n 입력하신 엑셀파일을 읽어오고 있습니다')
excel = pd.read_excel(upload_path, names=['사이트명', '사이트주소', '사용아이디', '업로드여부', '파일명'])
input_id_list = list(excel['사용아이디'].drop_duplicates())

n_error_list = []
d_error_list = []

def content_html():
    driver.switch_to.window(tabs[0])
    driver.get(content_path)
    ActionChains(driver).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
    ActionChains(driver).key_down(Keys.CONTROL).send_keys('c').key_up(Keys.CONTROL).perform()
def login():
    driver.switch_to.window(tabs[1])
    login_url = "https://nid.naver.com/nidlogin.login"
    driver.get(login_url)

    pyperclip.copy(auth)
    driver.find_element(By.ID, 'id').send_keys(Keys.CONTROL + 'v')
    pyperclip.copy(auth_dic[auth])
    driver.find_element(By.ID, 'pw').send_keys(Keys.CONTROL + 'v')
    login_btn = driver.find_element(By.ID, 'log.login')
    login_btn.click()
    print("로그인: 로그인 완료\n")
    time.sleep(2)
    a = input("2차 인증 여부 확인해주시고 아무거나 입력 후 엔터")


def posting():
    driver.switch_to.window(tabs[1])
    driver.get(naver_url)
    print("링크 접속 완료")
    time.sleep(1)

    # 포스팅
    driver.find_element(By.XPATH, '//a[contains(text(),"나의활동")]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//a[contains(text(),"내가 쓴 글 보기")]').click()
    time.sleep(3)

    # frame으로 변환해야 게시글 확인이 가능하다#
    driver.switch_to.frame("cafe_main")
    mcn = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/table/tbody/tr[1]/td[1]/div[3]/div/a')
    mcn_lnk = mcn.get_attribute('href')
    driver.get(mcn_lnk)
    time.sleep(2)

    driver.switch_to.frame("cafe_main")
    mcn = driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[3]/div[1]/a[1]')
    mcn_lnk = mcn.get_attribute('href')
    driver.get(mcn_lnk)
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="app"]/div/div/section/div/div[2]/div[1]/div[1]/div[2]/div/textarea').click()
    action = ActionChains(driver)
    action.send_keys(title).perform()
    print("글쓰기 1/3: 제목 입력 완료")

    time.sleep(1)
    content_html()
    driver.switch_to.window(tabs[1])
    time.sleep(1)
    driver.find_elements(By.XPATH, '//p[contains(@class,"se-text-paragraph se-text-paragraph-align-left")]')[0].click()
    ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()

    print("글쓰기 2/3: html 코드 입력 완료")

    time.sleep(1)
    driver.find_element(By.XPATH, '//span[contains(@class,"BaseButton__txt")]').click()
    print("글쓰기 3/3: 업로드 완료")
    time.sleep(3)


# 실행되는 라인
print('사용할 수 있는 아이디는 다음과 같습니다:' + str(auth_dic.keys()))
print('제공된 엑셀 참조 파일에 의하면 다음과 같은 id 가 제공되었습니다: ' + str(input_id_list))
global m
m = 0
List = [x for x in input('\n사용할 아이디를 알려주세요.(띄어쓰기로 구분합니다)\n').split()]
for x in List:
    auth = x
    print("사용할 아이디: " + auth)
    print("사용될 패스워드: " + auth_dic[auth])

    # 데이터 전처리
    excel_1 = excel[excel['사용아이디'] == list(auth_dic.keys())[m]]
    print(excel_1)
    url_list = list(excel_1['사이트주소'])
    naver_list = [x for x in url_list if "cafe.naver.com" in x]
    len_naver = len(naver_list)
    daum_list = [y for y in url_list if "cafe.daum.net" in y]
    len_daum = len(daum_list)
    m = +1

    login()
    print("\n" + auth + " 아이디로 네이버 카페 " + str(len_naver) + "개, 다음 카페 " + str(len_daum) + "개를 진행하겠습니다")

    i = 0
    while i <= len_naver:
        try:
            naver_url = naver_list[i]
            posting()
            excel_1.iloc[i, 3] = "O"
        except:
            if i >= len_naver:
                pass
            else:
                n_error_list.append(naver_list[i])
                excel_1.iloc[i, 3] = "X"
        i = i + 1
    naver_list = []
    daum_list = []

if len(n_error_list) > 0:
    print("다음은 권한이 없거나 오류가 있어서 업로드 하지 못한 링크들 입니다. 네이버 카페:")
    print(n_error_list)
else:
    pass
if len(d_error_list) > 0:
    print("다음은 권한이 없거나 오류가 있어서 업로드 하지 못한 링크들 입니다. 다음 카페:")
    print(d_error_list)
else:
    pass
excel_1.to_excel('C:\\Users\\movem\\Desktop\\' + auth + '_작업완료.xlsx')
print('작업완료된 내역을 엑셀 파일로 저장하였습니다.')
a = input()
