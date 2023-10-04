import pandas as pd
import os
# 엑셀 파일 경로
print("포스팅 전 파일 분배기입니다")
print("엑셀 파일 제시순서는 사이트명, 사이트주소, 업로드여부, 사용아이디, 파일명으로 지정해주십시오")
print("참조할 엑셀파일을 지정해주십시오")
excel_file = input().replace('"', '')
print("엑셀파일이 입력되었습니다")
# 데이터프레임 생성
df = pd.read_excel(excel_file, usecols=['사이트명', '사이트주소', '업로드여부', '사용아이디', '파일명'])
pd.set_option('mode.chained_assignment', None)
if len(df.loc[1, '사용아이디']) > 3:
    # 순서 변경된 데이터프레임 생성
    df[['업로드여부', '사용아이디']] = df[['사용아이디', '업로드여부']]
    print("포스팅 프로그램이 인식할 수 있도록 테이블 순서를 일부 변경합니다")
    print(df)

# "cafe.naver" 필터링 후 저장
naver_df_re = df[df['사이트주소'].str.contains('cafe.naver.com')]
naver_file_path = os.path.join(os.path.dirname(excel_file), "naver_cafe.xlsx")
naver_df_re.rename(columns={'사용아이디': '업로드여부', '업로드여부': '사용아이디'}, inplace=True)
naver_df_re.to_excel(naver_file_path, index=False)
print("네이버 완료")

# "cafe.daum.net" 필터링 후 저장
daum_df_re = df[df['사이트주소'].str.contains('cafe.daum.net')]
daum_file_path = os.path.join(os.path.dirname(excel_file), "daum_cafe.xlsx")
daum_df_re.rename(columns={'사용아이디': '업로드여부', '업로드여부': '사용아이디'}, inplace=True)
daum_df_re.to_excel(daum_file_path, index=False)
print("다음 완료")

# 나머지 경우 필터링 후 저장
other_df_re = df[~(df['사이트주소'].str.contains('cafe.naver.com') | df['사이트주소'].str.contains('cafe.daum.net'))]
other_file_path = os.path.join(os.path.dirname(excel_file), "other_cafe.xlsx")
other_df_re.rename(columns={'사용아이디': '업로드여부', '업로드여부': '사용아이디'}, inplace=True)
other_df_re .to_excel(other_file_path, index=False)
print("기타 완료\n")
print("완료되었습니다")
a = input()
