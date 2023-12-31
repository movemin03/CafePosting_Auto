from selenium import webdriver
from selenium.common import NoSuchElementException, NoSuchWindowException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import openpyxl
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import pyperclip
from selenium.webdriver.common.keys import Keys
import time
import re

# 사용자 지정..
ver = str("2024-01-06")
auth_dic = {'id': 'pw'}
chrome_ver = 120
filter_list = ['사이트명', '사이트주소', '사용아이디', '업로드여부', '파일명']

# 크롬 로딩..
login_status = 0
options = webdriver.ChromeOptions()
options.add_argument("headless")

driver = webdriver.Chrome(options=options)
#driver = webdriver.Chrome()

root = tk.Tk()
root.geometry("500x600")
root.title("네이버 카페 Tracker")

keyword_label = ttk.Label(root, text="검색할 공모전 명")
keyword_label.pack()
keyword_entry = ttk.Entry(root, width=67)
keyword_entry.pack()

excel_label = ttk.Label(root, text="공모전 목록 엑셀 위치")
excel_label.pack()
excel_entry = ttk.Entry(root, width=67)
excel_entry.pack()

options_label = ttk.Label(root, text="검색 필터 강도 선택")
options_label.pack()
filter_options = {
    "정확": 98,
    "표준": 90,
    "느슨함": 70,
    "해제": 0
}

selected_option = tk.StringVar(root)
selected_option.set("표준")

filter_value = filter_options[selected_option.get()]


def update_filter_value(*args):
    selected_menu = selected_option.get()
    global filter_value
    filter_value = filter_options[selected_menu]
    return filter_value


selected_option.trace("w", update_filter_value)
option_menu = tk.OptionMenu(root, selected_option, *filter_options.keys())
option_menu.pack()


def example_xlsx():
    ex_df = pd.DataFrame(columns=filter_list)
    user = os.getlogin()
    ex_df_time = datetime.today().strftime("%Y%m%d_%H%M%S")
    file_path = "C:\\Users\\" + user + "\\Desktop\\입력양식_네이버카페_tracker_" + str(ex_df_time) + '.xlsx'
    ex_df.to_excel(file_path, index=False)
    messagebox.showinfo(title="알림", message=file_path + ".xlsx 에 양식이 저장되었습니다")


keyword_label = ttk.Label(root, text="기타옵션")
keyword_label.pack()
button = ttk.Button(root, text="엑셀 입력 양식 생성", command=example_xlsx)
button.pack()

Result_Viewlabel = ttk.LabelFrame(text="실행내용")
Result_Viewlabel.pack()

Result_Viewlabel_Scrollbar = tk.Listbox(Result_Viewlabel,
                                        selectmode="extended", width=100, height=25, font=('Normal', 10), yscrollcommand=tk.Scrollbar(Result_Viewlabel).set)
Result_Viewlabel_Scrollbar.pack(side=tk.LEFT, fill=tk.BOTH)

pb_type = tk.DoubleVar()
progress_bar = ttk.Progressbar(orient="horizontal", length=400, mode="determinate", maximum=100, variable=pb_type)
progress_bar.pack()

progress = 0
execute_num = 0

data = {
            'post_key': [],
            'post_clubid': [],
            'post_member': [],
            'post_level': [],
            'orgin_link': [],
            'post_link': [],
            'post_title': [],
            'post_name': [],
            'orgin_id': [],
            'post_date': [],
            'post_view': [],
            'post_comments': [],
            'orgin_file_name': []
        }


def restart_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    # 크롬창 headless 모드를 해제하려면 아래 줄 주석 해제
    global driver
    # driver = webdriver.Chrome()
    driver = webdriver.Chrome(options=options)
    # 로딩이 완료될 때까지 최대 10초간 대기


def calculate_similarity(main_txt, sub_txt):
    str_to_vec = lambda s: [ord(c) for c in s]

    main_vec = str_to_vec(main_txt)
    sub_vec = str_to_vec(sub_txt)

    dot_product = sum(a * b for a, b in zip(main_vec, sub_vec))
    magnitude = (sum(a * a for a in main_vec) * sum(b * b for b in sub_vec)) ** 0.5

    return dot_product / magnitude


def work(url, keyword, user_id, or_file_name_item):
    if url.count("/") > 3:
        split_text = url.split("/")
        split_text = split_text[:4]  # 처음 4개의 요소만 유지
        url = "/".join(split_text)

    global login_status
    print("로그인 완료 여부:" + str(login_status))
    if login_status == 1:
        driver.get("https://www.naver.com/")
        session_cookies = visible_driver.get_cookies()
        for cookie in session_cookies:
            driver.add_cookie(cookie)
        driver.refresh()
        visible_driver.quit()
        login_status = 0

    driver.get(url)

    Result_Viewlabel_Scrollbar.insert(tk.END, url + " 카페로 접속합니다")
    Result_Viewlabel_Scrollbar.see(tk.END)
    progress = 10
    pb_type.set(progress)
    progress_bar.update()

    Result_Viewlabel_Scrollbar.update()

    global wait
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_element_located((By.XPATH, '//div[@id="info-search"]')))

    progress = 20
    pb_type.set(progress)
    progress_bar.update()
    try:
        member = driver.find_element(By.XPATH, '//li[@class="mem-cnt-info"]/child::*[2]')
        level = driver.find_element(By.XPATH, '//*[@id="cafe-info-data"]/div[1]/div[2]/ul/li[1]')
        member_txt = member.text.replace("비공개", "").replace("공개", "")
        level_txt = level.text.replace("비공개", "").replace("공개", "").replace("카페등급", "").replace("\n", "")
    except NoSuchElementException:
        Result_Viewlabel_Scrollbar.insert(tk.END, "멤버수와 카페등급 칸이 일반적 위치에 있지 않습니다")
        member_txt, level_txt = "n/a", "n/a"

    try:
        driver.find_element(By.XPATH, '//div[@id="info-search"]/child::*[1]/child::*[1]').click()
        search_box = driver.find_element(By.XPATH, '//div[@id="info-search"]/child::*[1]/child::*[1]')
        search_box.clear()
        search_box.send_keys(keyword)
        pyperclip.copy(keyword)
        driver.find_element(By.XPATH, '//div[@id="info-search"]/child::*[1]/child::*[2]').click()
    except NoSuchElementException:
        try:
            driver.find_element(By.XPATH, '//div[@id="cafe-search"]/child::*[1]/child::*[1]').click()
            search_box = driver.find_element(By.XPATH, '//div[@id="cafe-search"]/child::*[1]/child::*[1]')
            search_box.clear()  # 입력 필드 초기화
            search_box.send_keys(keyword)  # 검색어 입력
            driver.find_element(By.XPATH, '//div[@id="cafe-search"]/child::*[1]/child::*[2]').click()
        except NoSuchElementException:
            Result_Viewlabel_Scrollbar.insert(tk.END, "알 수 없는 오류로 페이지 로딩에 실패했습니다. 검색 버튼을 다시 눌러주세요")
    try:
        driver.switch_to.frame("cafe_main")
        wait.until(EC.presence_of_element_located((By.XPATH, '//div[@class="sort_area"]/child::*[3]/child::*[1]')))
        driver.find_element(By.XPATH, '//div[@class="sort_area"]/child::*[3]/child::*[1]').click()
        driver.find_element(By.XPATH, '//div[@class="sort_area"]/child::*[3]/child::*[2]/child::*[7]').click()
    except NoSuchElementException:
        print("50개씩 보기 단추 클릭 실패")

    progress = 50
    pb_type.set(progress)
    progress_bar.update()

    selected_option.trace("w", update_filter_value)
    rows_xpath = "//*[@id='main-area']/div[5]/table/tbody/tr"  # 모든 행을 한번에 선택
    rows = driver.find_elements(By.XPATH, rows_xpath)

    print("옵션 값" + str(filter_value))

    for i, row in enumerate(rows):
        try:
            try:
                posting_link_track = "./child::*[1]/child::*[2]"
                posting_link = row.find_element(By.XPATH,posting_link_track + "/child::*[1]/child::*[1]").get_attribute("href")
                #print("method1")
            except:
                try:
                    posting_link_track = "./child::*[1]/child::*[3]"
                    posting_link = row.find_element(By.XPATH, posting_link_track + "/child::*[1]/child::*[1]").get_attribute("href")
                    #print("method2")
                except:
                    if i == 0:
                        Result_Viewlabel_Scrollbar.insert(tk.END, url + "method3 가 필요한 카페입니다")
                    else:
                        Result_Viewlabel_Scrollbar.insert(tk.END, url + " 1번째 항목도 아닌데 왜 오류가 발생... method3 가 필요한 카페입니다")
            posting_cafe_clubid = re.search('clubid=(\\d+)', posting_link).group(1)
            posting_cafe_articleid = re.search('articleid=(\\d+)', posting_link).group(1)
            posting_key = str(posting_cafe_clubid) + str(posting_cafe_articleid)

            posting_title = row.find_element(By.XPATH, posting_link_track + "/child::*[1]/child::*[1]").text

            similarity = calculate_similarity(posting_title.replace(" ", ""), keyword.replace(" ", ""))

            if similarity * 100 > filter_value:
                Result_Viewlabel_Scrollbar.insert(tk.END, f'{i + 1}.게시물 "{posting_title}" 을 찾았습니다')
                Result_Viewlabel_Scrollbar.see(tk.END)
                try:
                    comments_pre = row.find_elements(By.XPATH, posting_link_track + "/child::*[1]/*")[1:]
                    comments_pre2 = [child.text for child in comments_pre]
                    posting_comments = ' '.join(comments_pre2).replace("[", "").replace("]", "").replace(" ", "").replace(
                        "사진", "").replace("파일", "").replace("링크", "").replace("new", "").replace("투표", "")
                    if posting_comments == "":
                        posting_comments = "0"
                except Exception as e:
                    print('comment가 없음.', e)
                    posting_comments = "n/a"

                posting_name = row.find_element(By.XPATH, "./child::*[2]").text
                posting_date = row.find_element(By.XPATH, "./child::*[3]").text
                posting_view = row.find_element(By.XPATH, "./child::*[4]").text

                user_id = re.sub(r'\d', '', user_id)
                user_id = user_id.replace(" ", "")

                data['post_key'].append(posting_key)
                data['post_clubid'].append(posting_cafe_clubid)
                data['post_member'].append(member_txt)
                data['post_level'].append(level_txt)

                data['orgin_link'].append(url)
                data['post_link'].append(posting_link)
                data['post_title'].append(posting_title)
                data['post_name'].append(posting_name)
                data['post_date'].append(posting_date)
                data['post_view'].append(posting_view)
                data['post_comments'].append(posting_comments)
                data['orgin_id'].append(user_id)
                data['orgin_file_name'].append(or_file_name_item)

            else:
                Result_Viewlabel_Scrollbar.insert(tk.END, f'검색어와 달라 제외: "{posting_title}"')
            Result_Viewlabel_Scrollbar.see(tk.END)
        except Exception as e:
            Result_Viewlabel_Scrollbar.insert(tk.END, f'{i + 1}번째 항목이 존재하지 않습니다:{url}')
            print(e)


    Result_Viewlabel_Scrollbar.insert(tk.END, url + " :검색이 완료되었습니다")
    Result_Viewlabel_Scrollbar.see(tk.END)
    progress = 100
    pb_type.set(progress)
    progress_bar.update()


def get_site_address_list(excel_input):
    excel = pd.read_excel(excel_input)
    pd.set_option('mode.chained_assignment', None)

    global input_id_list
    input_id_list = list(excel['사용아이디'].drop_duplicates())
    try:
        input_id_list = [item.strip() for item in input_id_list]
        input_id_list = [string.split("/")[0].strip() for string in input_id_list]
    except:
        Result_Viewlabel_Scrollbar.insert(tk.END, "제공하신 엑셀파일의 아이디 칸이 비어 있습니다. 프로그램을 재실행하시길 권장합니다")
        Result_Viewlabel_Scrollbar.see(tk.END)

    grouped = excel.groupby('사용아이디')
    dfs = {}
    for name, group in grouped:
        dfs[name] = group

    individual_dfs = {}
    for id in input_id_list:
        if id in dfs:
            individual_dfs[id] = dfs[id]

    return individual_dfs


def login(user_id):
    visible_options = webdriver.ChromeOptions()
    global visible_driver
    visible_driver = webdriver.Chrome(options=visible_options)
    global tabs
    tabs = visible_driver.window_handles
    login_url = "https://nid.naver.com/nidlogin.login"
    visible_driver.get(login_url)
    Result_Viewlabel_Scrollbar.insert(tk.END, "로그인 창에 접속중입니다")
    Result_Viewlabel_Scrollbar.update()

    action = ActionChains(visible_driver)

    visible_driver.find_element(By.ID, 'id').click()
    action.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
    action.send_keys(Keys.BACKSPACE).perform()
    pyperclip.copy(user_id)
    visible_driver.find_element(By.ID, 'id').send_keys(Keys.CONTROL + 'v')

    try:
        password = auth_dic[user_id].replace("\n", "").replace("", "")
    except:
        Result_Viewlabel_Scrollbar.insert(tk.END, "auth_dic 에 입력되어 있는 패스워드에 오류가 있습니다")
        password = input()

    visible_driver.find_element(By.ID, 'pw').click()
    action.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
    action.send_keys(Keys.BACKSPACE).perform()
    pyperclip.copy(password)
    visible_driver.find_element(By.ID, 'pw').send_keys(Keys.CONTROL + 'v')


    time.sleep(1)
    login_btn = visible_driver.find_element(By.ID, 'log.login')
    login_btn.click()
    Result_Viewlabel_Scrollbar.insert(tk.END, "로그인 작업 완료")
    Result_Viewlabel_Scrollbar.see(tk.END)

    try:
        if "https://nid.naver.com/login/ext/deviceConfirm" in visible_driver.current_url:
            wait_s = WebDriverWait(visible_driver, 1)
            wait_s.until(EC.presence_of_element_located((By.XPATH, '//*[@id="new.save"]')))
            visible_driver.find_element(By.XPATH, '//*[@id="new.save"]').click()
    except:
        pass
    while True:
        try:
            wait_s = WebDriverWait(visible_driver, 60)
            wait_s.until(EC.presence_of_element_located((By.XPATH, '//*[@id="account"]/div[1]/div/div/div[2]')))
            id_chk = visible_driver.find_element(By.XPATH, '//*[@id="account"]/div[1]/div/div/div[2]').text.replace(
                "@naver.com", "")
            global manual_not_login
            manual_not_login = 0
            if user_id != id_chk:
                print(user_id + ":" + id_chk)
                Result_Viewlabel_Scrollbar.insert(tk.END, "2차 인증 여부 및 아이디가 " + user_id + "가 맞는지 확인해주시길")
                Result_Viewlabel_Scrollbar.see(tk.END)
                manual_not_login = 1
            break
        except Exception as e:
            print(e)
            answer = messagebox.askyesno("60초 타임아웃", "2차 인증이 완료되지 않았습니다. 재시도할까요?")
            if answer:
                Result_Viewlabel_Scrollbar.insert(tk.END, "인증을 재시도합니다")
                Result_Viewlabel_Scrollbar.see(tk.END)
            else:
                Result_Viewlabel_Scrollbar.insert(tk.END, "로그인하지 않습니다")
                Result_Viewlabel_Scrollbar.see(tk.END)
                manual_not_login = 1
                break
        global login_status
    login_status = 1
    return manual_not_login


def result():
    progress = 0
    pb_type.set(progress)
    progress_bar.update()

    # 데이터프레임 생성
    df = pd.DataFrame(data)
    duplicated_data = df[df.duplicated(subset='post_link', keep=False)]
    df = df.drop_duplicates(subset='post_link')


    file_time = datetime.today().strftime("%Y%m%d_%H%M")
    user = os.getlogin()
    file_path = 'C:\\Users\\' + user + '\\Desktop\\cafe_data_' + keyword + "_" + file_time
    file_path_dupicated = 'C:\\Users\\' + user + '\\Desktop\\cafe_data_중복된데이터_' + keyword + "_" + file_time

    # 엑셀 파일로 저장
    with pd.ExcelWriter(file_path + ".xlsx") as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        Result_Viewlabel_Scrollbar.insert(tk.END, "\n검색된 데이터가" + file_path + ".xlsx 에 저장 완료되었습니다")

    if not duplicated_data.empty:
        with pd.ExcelWriter(file_path_dupicated + ".xlsx") as writer:
            duplicated_data.to_excel(writer, index=False, sheet_name='Sheet1')
        Result_Viewlabel_Scrollbar.insert(tk.END, "\n중복 데이터가" + file_path_dupicated + ".xlsx 에 저장 완료되었습니다")
        Result_Viewlabel_Scrollbar.insert(tk.END, "중복 데이터는 사용자가 제공한 엑셀 파일에 다른 아이디로 동일한 카페에 접속하도록 적어둔 경우 발생")

    Result_Viewlabel_Scrollbar.see(tk.END)

    pb_type.set(100)
    progress_bar.update()
    messagebox.showinfo(title="알림", message="검색이 완료되어 " + file_path + ".xlsx 에 저장 완료되었습니다")

def execute():
    global keyword, excel_input
    keyword = keyword_entry.get()
    excel_input = excel_entry.get().replace("'", "").replace('"', "").replace("\\", "\\")
    if keyword == "" or excel_input == "":
        messagebox.showinfo(title="알림", message="검색할 공모전 이름와 엑셀파일 위치가 입력되지 않았습니다. 둘 다 입력해주십시오")
    else:
        Result_Viewlabel_Scrollbar.insert(tk.END, "검색할 공모전 이름과 엑셀파일이 입력되었습니다")
        Result_Viewlabel_Scrollbar.insert(tk.END, "엑셀 파일을 프로그램이 읽을 수 있는 형태로 변환하는 중입니다")
        Result_Viewlabel_Scrollbar.update()
        start_time = datetime.now()
        Result_Viewlabel_Scrollbar.update()

        try:
            individual_dfs = get_site_address_list(excel_input)
            Result_Viewlabel_Scrollbar.insert(tk.END, "검색을 위한 크롬드라이버를 로딩하고 있는 중입니다")
            Result_Viewlabel_Scrollbar.update()
        except FileNotFoundError:
            Result_Viewlabel_Scrollbar.insert(tk.END, "잘못된 엑셀 경로가 입력되었습니다")
        except KeyError:
            Result_Viewlabel_Scrollbar.insert(tk.END, "잘못된 엑셀 데이터가 입력되었습니다")
        except ValueError:
            Result_Viewlabel_Scrollbar.insert(tk.END, "잘못된 엑셀 데이터가 입력되었습니다: 엑셀 파일이 아닙니다")
        try:
            for user_id, df in individual_dfs.items():
                print(f"사용자 아이디: {user_id}")
                print(f"데이터 프레임: \n{df}\n")
                list_individual = df['사이트주소'].tolist()
                login(user_id)
                for url in list_individual:
                    try:
                        if manual_not_login == 0:
                            try:
                                matching_record = df[df['사이트주소'] == url].iloc[0]
                                or_file_name_item = matching_record['파일명']
                                if len(or_file_name_item) == 0:
                                    or_file_name_item = "n/a"
                            except Exception as e:
                                print("파일명(커뮤니티명) 가져오기 오류:")
                                print(e)
                            work(url, keyword, user_id, or_file_name_item)
                    except NoSuchWindowException:
                        restart_driver()
            Result_Viewlabel_Scrollbar.insert(tk.END, "검색이 완료되어 결과를 내보내는 중입니다")
            Result_Viewlabel_Scrollbar.update()
            end_time = datetime.now()
            execution_time = end_time - start_time
            result()
            Result_Viewlabel_Scrollbar.insert(tk.END, f"검색 소요 시간: {execution_time}")
        except UnboundLocalError as e:
            print(e)
            Result_Viewlabel_Scrollbar.insert(tk.END, "잘못된 엑셀 양식이 입력되었습니다")
            Result_Viewlabel_Scrollbar.insert(tk.END, "사이트명, 사이트주소, 사용아이디, 업로드여부, 파일명 순이어야 합니다")
        except ValueError as e:
            print("원치 않은 오류:")
            print(e)
            list_lengths = {key: len(value) for key, value in data.items()}
            max_length = max(list_lengths.values())
            different_lengths = {key: value for key, value in list_lengths.items() if value != max_length}
            print("항목 개수가 다른 리스트들:")
            for key, value in different_lengths.items():
                print(f"{key}: {value} 개")


button = ttk.Button(root, text="검색", command=execute)
button.pack()

root.mainloop()
