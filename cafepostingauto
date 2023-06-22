from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
import time
import pyperclip
import pandas as pd

ver = str("2023-06-21")

# 안내
print("\n")
print("씽굿 프로그램입니다.")
print("https://github.com/movemin03/CafePosting_Auto")
print("ver:" + ver)

# 정보 입력
print('필요한 정보를 기입해야합니다\n')
print("게시물의 제목을 적어주십시오:")
title = "[추천공모전]예시"
print("복사할 html 코드입력해주세요:")
cnt = "html예시"

# 사용할 아이디/패스워드
auth_dic = {'example_id': 'example_pw'}
print("사용할 아이디: ")
auth = str('movemin03')
print("아이디: " + auth)
print("사용될 패스워드: " + auth_dic[auth])

# 데이터 전처리0
print('참고할 엑셀 파일 위치를 알려주세요:')
upload_path = "C:\\Users\\movem\\Desktop\\워크시트.xlsx"
excel = pd.read_excel(upload_path, names = ['사이트명', '사이트주소', '사용아이디', '업로드여부'])
excel_1 = excel[excel['사용아이디'] == auth]
url_list = list(excel_1['사이트주소'])

# 데이터 전처리1
print("\n 데이터 전처리 중... 네이버 카페와 다음 카페 분류 중 입니다\n")
naver_list = [x for x in url_list if "cafe.naver.com" in x]
len_naver = len(naver_list)
daum_list = [y for y in url_list if "cafe.daum.net" in y]
len_daum = len(daum_list)
n_error_list = []
d_error_list = []

def login():
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
    driver.find_elements(By.XPATH, '//p[contains(@class,"se-text-paragraph se-text-paragraph-align-left")]')[0].click()
    action = ActionChains(driver)
    action.send_keys(cnt).perform()
    print("글쓰기 2/3: html 코드 입력 완료")

    time.sleep(1)
    driver.find_element(By.XPATH, '//span[contains(@class,"BaseButton__txt")]').click()
    print("글쓰기 3/3: 업로드 완료")
    time.sleep(3)


# 실행되는 라인
driver = webdriver.Chrome()
login()
i = 0
while i <= len_naver:
    try:
        naver_url = naver_list[i]
        posting()
        excel_1.iloc[i,3] = "O"
    except:
        if i >= len_naver:
            pass
        else:
            n_error_list.append(naver_list[i])
            excel_1.iloc[i, 3] = "X"
    i = i + 1
print("\n네이버 카페 " + str(len_naver) + "개, 다음 카페 " + str(len_daum) + "개를 시도하셨습니다.")

if len(n_error_list)>0:
    print("다음은 권한이 없거나 오류가 있어서 업로드 하지 못한 링크들 입니다. 네이버 카페:")
    print(n_error_list)
else:
    pass
if len(d_error_list)>0:
    print("다음은 권한이 없거나 오류가 있어서 업로드 하지 못한 링크들 입니다. 다음 카페:")
    print(d_error_list)
else:
    pass
excel_1.to_excel('C:\\Users\\movem\\Desktop\\작업완료.xlsx')
print('작업완료 된 내역을 엑셀파일로 저장하였습니다.')
a = input()
