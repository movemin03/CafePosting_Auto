import os
import re
import pandas as pd

def separate_numbers_and_strings(folder_path):
    folder_name = []
    distributed = []
    completed = []
    star = []

    with os.scandir(folder_path) as entries:
        for entry in entries:
            if entry.is_dir():
                dirname = entry.name
                lines = dirname.split('\n')
                pattern1 = r"(\D+)\s(\d+)\s\((\d+)\)"
                pattern2 = r"(\D+)\s(\d+)"

                for line in lines:
                    try:
                        korean_pattern = r"[가-힣a-zA-Z]+"
                        non_korean_pattern = r"[^가-힣a-zA-Z]+"

                        korean_part = "".join(re.findall(korean_pattern, line))
                        non_korean_part = "".join(re.findall(non_korean_pattern, line))

                        if "★" in non_korean_part:
                            star_num = "★"
                            non_korean_part = non_korean_part.replace("★", "")
                        else:
                            star_num = ""

                        korean_part = korean_part.replace(' ', '') + ' '
                        non_korean_part = non_korean_part.replace(' ', '').replace('(', ' (')

                        line = str(korean_part) + str(non_korean_part)

                    except:
                        pass

                    match1 = re.search(pattern1, line)
                    match2 = re.search(pattern2, line)

                    if match1:
                        non_numbers = match1.group(1)
                        dis_numbers = match1.group(2)
                        com_numbers = match1.group(3)

                        folder_name.append(non_numbers)
                        distributed.append(dis_numbers)
                        completed.append(com_numbers)

                    if match2 and not match1:
                        non_numbers = match2.group(1)
                        dis_numbers = match2.group(2)

                        folder_name.append(non_numbers)
                        distributed.append(dis_numbers)
                        completed.append(0)

                    star.append(star_num)

    return star, folder_name, distributed, completed


def version_1():
    print("version_1: 엑셀파일로 출력")
    folder_path = input("폴더 경로를 입력하세요: ").replace('"', '')
    star, folder_name, distributed, completed = separate_numbers_and_strings(folder_path)

    result_dataframe = make_dataframe(star, folder_name, distributed, completed)
    print(result_dataframe)

    save_dataframe_to_excel(result_dataframe, folder_path)

def make_dataframe(star, folder_name, distributed, completed):
    folder_data = []
    for star_num, name, dis_num, com_num in zip(star, folder_name, distributed, completed):
        folder_data.append({"마크": star_num, "폴더 명": name, "분배": dis_num, "완료": com_num})
    return pd.DataFrame(folder_data)

def save_dataframe_to_excel(dataframe, folder_path, file_name='카테고리정리_output.xlsx'):
    output_file_path = os.path.join(folder_path, file_name)
    dataframe.to_excel(output_file_path, index=False)
    print(f"엑셀 파일이 저장되었습니다: {output_file_path}")
    print("종료하려면 아무키나 눌러주세요")


def version_2():
    print("version_2: 텍스트파일로 출력")
    folder_path = input("폴더 경로를 입력하세요: ").replace('"', '')
    star, folder_name, distributed, completed = separate_numbers_and_strings(folder_path)
    print("예시: 국가생물다양성전략 정책 아이디어 공모전(홍길동선임)")
    print("제목을 입력해주십시오:")
    global arrange_title
    arrange_title = input()
    save_formatted_txt(star, folder_name, distributed, completed, folder_path + "\\카테고리정리_파일.txt")
    print("파일로 저장되었습니다")

def save_formatted_txt(star, folder_name, distributed, completed, output_file):
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("*" + arrange_title)
        total_distributed = sum(map(int, distributed))
        total_completed = sum(map(int, completed))
        total_output = f"\n      : ★금일완료{total_completed}개, (총 {total_distributed}개 분배, 누적완료 {total_completed}개, 잔여 0개)\n"
        f.write(total_output)

        for star_num, name, dis_num, com_num in zip(star, folder_name, distributed, completed):
            left = str(int(dis_num) - int(com_num))
            output = f"        → {star_num}{name}(총 {dis_num}개 분배, 금일 {com_num}개, 잔여 {left}개)\n"
            f.write(output)

if __name__ == "__main__":
    print("포스팅 완료 후 작업일지 작성을 돕기 위한 프로그램입니다. ")
    print("버전을 선택하세요(숫자입력): 1. 엑셀 표로 출력, 2. 텍스트 파일로 출력 ")
    qa = input()
    if qa == "1":
        version_1()
    else:
        version_2()
    a = input()
