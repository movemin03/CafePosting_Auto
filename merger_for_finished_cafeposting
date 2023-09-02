import glob
import os
from openpyxl import load_workbook
import pandas as pd

user = "movem"
filter_list = ['커뮤니티 명', 'URL주소', '아이디/비밀번호', 'O/X']
filter_list_2 = ['사이트명', '사이트주소', '사용아이디', '업로드여부', '파일명']

def combined_first():
    print('파일 병합을 하시겠습니까? y/n')
    a = input()
    if a == "y":
        combine_path = "C:\\Users\\" + user + "\\Desktop\\취합"
        print("병합된 파일은 여기에 저장됩니다: " + combine_path)
        print('위치 변경을 원하시나요? 변경 원한다면 y 입력')
        a = input()
        if a == "y":
            combine_path = input().replace('"', '')
            combine_path = combine_path + "\\"
        else:
            combine_path = combine_path + "\\"

        target = (str(combine_path) + '*.xlsx')
        print(target)
        xlsx_list = glob.glob(target)

        print("현재 있는 파일 :" + str(xlsx_list))

        dfs = []
        try:
        # 각 xlsx 파일에 대해 반복합니다.
            for xlsx_file in xlsx_list:
            # 엑셀 파일 읽기
                wb = load_workbook(xlsx_file)
                ws = wb.active
                row_heights = {}

                for row in range(1, ws.max_row + 1):
                    if row in ws.row_dimensions:
                        row_heights[row] = ws.row_dimensions[row].height
                    else:
                        row_heights[row] = "h_no"
                none_value_keys = [key for key, value in row_heights.items() if value is None]
                print("숨김 행:" + str(none_value_keys))

                excel = pd.read_excel(xlsx_file, header=0, skiprows=lambda x: x+1 in none_value_keys)
                excel_re = excel[filter_list]
                file_name = xlsx_file[len(combine_path):]
                print(file_name)
                file_name = file_name.replace(".xlsx", "")
                file_name_values = excel_re.apply(lambda row: file_name + "-"+str(row.name + 1), axis=1)
                excel_re["파일명"] = file_name_values
                excel_re.fillna('', inplace=True)

        # 데이터프레임 리스트에 추가합니다.
                dfs.append(excel_re)

            # 데이터프레임을 모두 결합합니다.
                combined_df = pd.concat(dfs, ignore_index=True)

            # 새 엑셀 파일 저장하기
                output_file = os.path.join(combine_path, "combined.xlsx")
                combined_df.to_excel(output_file, index=False, header=True)
                print(f"합쳐진 파일이 {output_file}에 저장되었습니다.")
                error_code = 0
        except:
            print("형식이 다른 파일이 끼어있습니다. 찾아서 삭제해주세요. 다시 시도합니다: ")
            error_code = 1
    else:
        pass
def combined():
    print('파일 병합을 하시겠습니까? y/n')
    a = input()
    if a == "y":
        combine_path = "C:\\Users\\" + user + "\\Desktop\\결과"
        print("병합된 파일은 여기에 저장됩니다: " + combine_path)
        print('위치 변경을 원하시나요? 변경 원한다면 y 입력')
        a = input()
        if a == "y":
            combine_path = input().replace('"', '')
            combine_path = combine_path + "\\"
        else:
            combine_path = combine_path + "\\"

        target = (str(combine_path) + '*.xlsx')
        print(target)
        xlsx_list = glob.glob(target)

        print("현재 있는 파일 :" + str(xlsx_list))

        dfs = []
        try:
        # 각 xlsx 파일에 대해 반복합니다.
            for xlsx_file in xlsx_list:
            # 엑셀 파일 읽기
                wb = load_workbook(xlsx_file)
                ws = wb.active
                row_heights = {}

                for row in range(1, ws.max_row + 1):
                    if row in ws.row_dimensions:
                        row_heights[row] = ws.row_dimensions[row].height
                    else:
                        row_heights[row] = "h_no"
                none_value_keys = [key for key, value in row_heights.items() if value is None]
                print("숨김 행:" + str(none_value_keys))

                excel = pd.read_excel(xlsx_file, header=0, skiprows=lambda x: x+1 in none_value_keys)
                print(excel)
                excel_re = excel[filter_list_2]
                print(excel_re)
            # 데이터프레임 리스트에 추가합니다.
                dfs.append(excel_re)

            # 데이터프레임을 모두 결합합니다.
                combined_df = pd.concat(dfs, ignore_index=True)

            # 새 엑셀 파일 저장하기
                output_file = os.path.join(combine_path, "combined.xlsx")
                combined_df.to_excel(output_file, index=False, header=True)
                print(f"합쳐진 파일이 {output_file}에 저장되었습니다.")
                error_code = 0
        except:
            print("형식이 다른 파일이 끼어있습니다. 찾아서 삭제해주세요. 다시 시도합니다: ")
            error_code = 1
    else:
        pass

print("어느작업을 원하십니까? 1: 처음 파일 병합 2: 결과 파일 병합")
a = input()
if a == "1":
    combined_first()
else:
    combined()
