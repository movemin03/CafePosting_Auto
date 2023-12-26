from selenium import webdriver
from selenium.common import NoSuchElementException, NoSuchWindowException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import openpyxl
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import pyperclip
from selenium.webdriver.common.keys import Keys


#크롬 로딩..
options = webdriver.ChromeOptions()
options.add_argument("headless")
# driver = webdriver.Chrome() 크롬창 headless 모드를 해제하려면 이것 사용
global driver
driver = webdriver.Chrome()
#driver = webdriver.Chrome(options=options)

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


Result_Viewlabel = ttk.LabelFrame(text="실행내용")
Result_Viewlabel.pack()

Result_Viewlabel_Scrollbar = tk.Listbox(Result_Viewlabel, selectmode="extended", width=100, height=25, font=('Normal', 10), yscrollcommand=tk.Scrollbar(Result_Viewlabel).set)
Result_Viewlabel_Scrollbar.pack(side=tk.LEFT, fill=tk.BOTH)

pb_type = tk.DoubleVar()
progress_bar = ttk.Progressbar(orient="horizontal", length=400, mode="determinate", maximum=100,variable=pb_type)
progress_bar.pack()

progress = 0
execute_num = 0

data = {
            'pl_link': [],
            'pl_title': [],
            'pl_name': [],
            'pl_date': [],
            'pl_view': [],
            'pl_comments': []
        }
def restart_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    # 크롬창 headless 모드를 해제하려면 아래 줄 주석 해제
    # driver = webdriver.Chrome()
    driver = webdriver.Chrome()
    # driver = webdriver.Chrome(options=options)
    # 로딩이 완료될 때까지 최대 10초간 대기

def work(url, keyword):
    if url.count("/") > 3:
        split_text = url.split("/")
        split_text = split_text[:4]  # 처음 4개의 요소만 유지
        url = "/".join(split_text)
    Result_Viewlabel_Scrollbar.insert(tk.END, url + " 카페로 접속합니다")

    Result_Viewlabel_Scrollbar.see(tk.END)
    progress = 10
    pb_type.set(progress)
    progress_bar.update()

    Result_Viewlabel_Scrollbar.update()
    driver.get(url)

    global wait
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_element_located((By.XPATH, '//div[@id="info-search"]')))

    progress = 20
    pb_type.set(progress)
    progress_bar.update()

    try:
        driver.find_element(By.XPATH,'//div[@id="info-search"]/child::*[1]/child::*[1]').click()
        pyperclip.copy(keyword)
        driver.find_element(By.XPATH,'//div[@id="info-search"]/child::*[1]/child::*[1]').send_keys(Keys.CONTROL + 'v')
        driver.find_element(By.XPATH, '//div[@id="info-search"]/child::*[1]/child::*[2]').click()
    except NoSuchElementException:
        try:
            driver.find_element(By.XPATH, '//div[@id="cafe-search"]/child::*[1]/child::*[1]').click()
            pyperclip.copy(keyword)
            driver.find_element(By.XPATH, '//div[@id="cafe-search"]/child::*[1]/child::*[1]').send_keys(
                Keys.CONTROL + 'v')
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

    rows_xpath = "//*[@id='main-area']/div[5]/table/tbody/tr"  # 모든 행을 한번에 선택
    rows = driver.find_elements(By.XPATH, rows_xpath)

    for i, row in enumerate(rows):
        try:
            posting_link = row.find_element(By.XPATH,
                                            "./child::*[1]/child::*[2]/child::*[1]/child::*[1]").get_attribute("href")
            posting_title = row.find_element(By.XPATH, "./child::*[1]/child::*[2]/child::*[1]/child::*[1]").text

            try:
                comments_pre = row.find_elements(By.XPATH, "./child::*[1]/child::*[2]/child::*[1]/*")[1:]
                comments_pre2 = [child.text for child in comments_pre]
                posting_comments = ' '.join(comments_pre2).replace("[", "").replace("]", "").replace(" ", "").replace(
                    "사진", "").replace("파일", "").replace("링크", "")
                if posting_comments == "":
                    posting_comments = "0"
            except Exception as e:
                print('예외가 발생했습니다.', e)
                posting_comments = "n/a"

            posting_name = row.find_element(By.XPATH, "./child::*[2]").text
            posting_date = row.find_element(By.XPATH, "./child::*[3]").text
            posting_view = row.find_element(By.XPATH, "./child::*[4]").text

            data['pl_link'].append(posting_link)
            data['pl_title'].append(posting_title)
            data['pl_comments'].append(posting_comments)
            data['pl_name'].append(posting_name)
            data['pl_date'].append(posting_date)
            data['pl_view'].append(posting_view)

        except Exception as e:
            print(f'{i + 1}번째 항목이 존재하지 않습니다')

    Result_Viewlabel_Scrollbar.insert(tk.END, url + " :검색이 완료되었습니다")
    Result_Viewlabel_Scrollbar.see(tk.END)
    progress = 100
    pb_type.set(progress)
    progress_bar.update()

def get_site_address_list(excel_input):
    df = pd.read_excel(excel_input)
    global site_address_list
    site_address_list = df["사이트주소"].tolist()
    return site_address_list

def result():
    progress = 0
    pb_type.set(progress)
    progress_bar.update()

    # 데이터프레임 생성
    df = pd.DataFrame(data)

    file_time = datetime.today().strftime("%Y%m%d_%H%M")
    user = os.getlogin()
    file_path = 'C:\\Users\\' + user + '\\Desktop\\cafe_data_' + keyword + "_" + file_time

    # 엑셀 파일로 저장
    with pd.ExcelWriter(file_path + ".xlsx") as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

    Result_Viewlabel_Scrollbar.insert(tk.END, "\n" + file_path + ".xlsx 에 저장 완료되었습니다")
    Result_Viewlabel_Scrollbar.see(tk.END)

    pb_type.set(100)
    progress_bar.update()
    messagebox.showinfo(title="알림", message="검색이 완료되어 " + file_path + ".xlsx 에 저장 완료되었습니다")

def execute():
    global keyword, excel_input
    keyword = keyword_entry.get()
    excel_input = excel_entry.get().replace(" ", "").replace("'", "").replace('"', "").replace("new", "")
    if keyword == "" or excel_input == "":
        messagebox.showinfo(title="알림", message="검색할 공모전 이름와 엑셀파일 위치가 입력되지 않았습니다. 둘 다 입력해주십시오")
    else:
        Result_Viewlabel_Scrollbar.insert(tk.END, "검색할 공모전 이름과 엑셀파일이 입력되었습니다")
        Result_Viewlabel_Scrollbar.insert(tk.END, "엑셀 파일을 프로그램이 읽을 수 있는 형태로 변환합니다")
        Result_Viewlabel_Scrollbar.update()
        get_site_address_list(excel_input)
        Result_Viewlabel_Scrollbar.insert(tk.END, "제공된 url 에 접속하여 정보를 가져오겠습니다")
        start_time = datetime.now()
        Result_Viewlabel_Scrollbar.update()

        for url in site_address_list:
            try:
                work(url, keyword)
            except NoSuchWindowException:
                restart_driver()

        Result_Viewlabel_Scrollbar.insert(tk.END, "검색이 완료되어 결과를 내보내는 중입니다")
        Result_Viewlabel_Scrollbar.update()
        end_time = datetime.now()
        execution_time = end_time - start_time
        result()
        Result_Viewlabel_Scrollbar.insert(tk.END, f"검색 소요 시간: {execution_time}")

button = ttk.Button(root, text="검색", command=execute)
button.pack()

root.mainloop()
