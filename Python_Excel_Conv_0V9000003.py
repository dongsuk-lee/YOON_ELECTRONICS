#!/usr/bin/env python
# coding: utf-8

# ### pandas read

# In[1]:


import pandas as pd
import numpy as np
import openpyxl as xl
import os
#import os.path
#from tkinter import filedialog
#from tkinter import *
#from pathlib import Path


# ### 파일 읽기
# * os.path.basename(filename) - 파일명만 추출
# * os.path.dirname(filename) - 디렉토리 경로 추출
# * os.path.split(filename) - 경로와 파일명을 분리
# * os.path.splitdrive(filename) - 드라이브명과 나머지 분리 (MS Windows의 경우)
# * os.path.splitext(filename) - 확장자와 나머지 분리

# def file_open():
#     root = Tk()  # 루트 윈도우 생성
#     root.title("File Select...")
#     root.geometry("200x100")
#     root.resizable(0,0)
# 
#     files = filedialog.askopenfilenames(initialdir="./",\
#         title = "파일을 선택 해 주세요",\
#         filetypes = (("*.xlsx","*xlsx"),("*.xls","*xls"),("*.csv","*csv")))
#     #files 변수에 선택 파일 경로 넣기
# 
#     if files == '':
#         filedialog.showwarning("경고", "파일을 추가 하세요")    #파일 선택 안했을 때 메세지 출력
#     root.withdraw()  # 루트 윈도우 숨기기 
#     str_file = ''.join(files)
#     return str_file
# 

# ### 메세지

# In[2]:


print("다음과 같이 파일을 확인한 후 아무키나 눌러 주세요.\n")
print("거래처별매입매출장.xlsx\n")
print("거래처등록.xlsx\n")
#os.system("pause")


# ### 거래처 매입매출 자료
# * read
# * 거레처코드로 sort

# In[3]:


df_tax = pd.read_excel('거래처별매입매출장.xlsx', usecols=['일자', '전표번호', '코드', '거래처', '유형', '품목', '수량', '단가', '공급가액', '부가세', '합계'])
#df_tax = pd.read_excel(file_open(), usecols=['일자', '전표번호', '코드', '거래처', '유형', '품목', '수량', '단가', '공급가액', '부가세', '합계'])
#df_tax_total = df_tax['일자'].count()
# sort
#df_tax_sort = df_tax.sort_values(by = '전표번호')
#df_tax_sort


# In[4]:


df_tax['유형'].replace('면건', np.NaN, inplace = True)
df_tax.dropna(subset=['유형'], inplace=True)
df_tax_total = df_tax['일자'].count()


# ### 업체 정보 read

# In[5]:


df_account = pd.read_excel('거래처등록.xlsx', header = 14, usecols=['거래처코드', '거래처명', '사업자등록번호', '대표자성명', '업태', '종목', '주소', 'EMail'])
df_account_count = df_account['거래처명'].count()
df_account.rename(columns={'거래처코드':'코드'}, inplace=True)


# ### Data Merge
# * tax + account

# 

# In[6]:


df_m = pd.merge(df_account, df_tax, on="코드", how="right")  # 이름이 같아야지만 merge가 가능

#df_m['거래처명'] = df_m['거래처명'].str.replace(pat=r'[^\w]', repl=r'', regex=True)
df_m['일자'] = df_m['일자'].str.replace(pat=r'[^\w]', repl=r'', regex=True)
df_m['거래처명'] = df_m['거래처명'].str.split('/').str.get(0)


# ### 전표 신규 데이터 read

# In[7]:


df_input = pd.read_excel("양식.xlsx", header = 5)
#df_input.loc[0] = ''
#df_input['전자(세금)계산서 종류'].iloc[0] = '01'


# ### 전표 입력

# In[8]:


def TAX_NewDataInput(line, sumline):
    df_input.loc[line + 1] = ''
    
    df_input['전자(세금)계산서 종류'].iloc[sumline] = '01'

    df_input['공급자 등록번호'].iloc[sumline] = '4098654575'
    df_input['공급자 상호'].iloc[sumline] = '주식회사윤전자'
    df_input['공급자 성명'].iloc[sumline] = '이창규'
    df_input['공급자 사업장주소'].iloc[sumline] = '경기도 군포시 공단로140번길 52 1413호'
    df_input['공급자 업태'].iloc[sumline] = '제조업'
    df_input['공급자 종목'].iloc[sumline] = '발전기 및 전동기 수리업'
    df_input['공급자 이메일'].iloc[sumline] = 'yoon.leecg@gmail.com'
    df_input['영수(01),청구(02)'].iloc[sumline] = '02'

    df_input['작성일자'].iloc[sumline] = df_m['일자'].loc[line]
    df_input['공급받는자 등록번호'].iloc[sumline] = df_m['사업자등록번호'].loc[line]
    df_input['공급받는자 상호'].iloc[sumline] = df_m['거래처명'].loc[line]
    df_input['공급받는자 성명'].iloc[sumline] = df_m['대표자성명'].loc[line]
    df_input['공급받는자 사업장주소'].iloc[sumline] = df_m['업태'].loc[line]
    df_input['공급받는자 업태'].iloc[sumline] = df_m['종목'].loc[line]
    df_input['공급받는자 종목'].iloc[sumline] = df_m['주소'].loc[line]
    df_input['공급받는자 이메일1'].iloc[sumline] = df_m['EMail'].loc[line]


# In[9]:


# 일자1	품목1	규격1	수량1	단가1	공급가액1	세액1
total_sum_price = 0
total_sum_tax = 0

def TAX_ReleaseDataInput(line, add, sumline):
    if(add == 0):
        df_input['일자1'].iloc[sumline]         = df_m['일자'].str.slice(8, 10).loc[line]
        df_input['품목1'].iloc[sumline]         = df_m['품목'].loc[line]
        df_input['수량1'].iloc[sumline]         = df_m['수량'].loc[line]
        df_input['단가1'].iloc[sumline]         = df_m['단가'].loc[line]
        df_input['공급가액1'].iloc[sumline]     = df_m['공급가액'].loc[line]
        df_input['세액1'].iloc[sumline]         = df_m['부가세'].loc[line]
        df_input['공급가액'].iloc[sumline]      = df_m['공급가액'].loc[line]
        df_input['세액'].iloc[sumline]      = df_m['부가세'].loc[line]

    elif(add == 1):
        df_input['일자2'].iloc[sumline]        = df_m['일자'].str.slice(8, 10).loc[line]
        df_input['품목2'].iloc[sumline]        = df_m['품목'].loc[line]
        df_input['수량2'].iloc[sumline]        = df_m['수량'].loc[line]
        df_input['단가2'].iloc[sumline]        = df_m['단가'].loc[line]
        df_input['공급가액2'].iloc[sumline]    = df_m['공급가액'].loc[line]
        df_input['세액2'].iloc[sumline]        = df_m['부가세'].loc[line]
        df_input['공급가액'].iloc[sumline]      = df_input['공급가액'].iloc[sumline] + df_m['공급가액'].loc[line]
        df_input['세액'].iloc[sumline]      = df_input['세액'].iloc[sumline] + df_m['부가세'].loc[line]

    elif(add == 2):
        df_input['일자3'].iloc[sumline]        = df_m['일자'].str.slice(8, 10).loc[line]
        df_input['품목3'].iloc[sumline]        = df_m['품목'].loc[line]
        df_input['수량3'].iloc[sumline]        = df_m['수량'].loc[line]
        df_input['단가3'].iloc[sumline]        = df_m['단가'].loc[line]
        df_input['공급가액3'].iloc[sumline]    = df_m['공급가액'].loc[line]
        df_input['세액3'].iloc[sumline]        = df_m['부가세'].loc[line]
        df_input['공급가액'].iloc[sumline]      = df_input['공급가액'].iloc[sumline] + df_m['공급가액'].loc[line]
        df_input['세액'].iloc[sumline]      = df_input['세액'].iloc[sumline] + df_m['부가세'].loc[line]

    elif(add == 3):
        df_input['일자4'].iloc[sumline]        = df_m['일자'].str.slice(8, 10).loc[line]
        df_input['품목4'].iloc[sumline]        = df_m['품목'].loc[line]
        df_input['수량4'].iloc[sumline]        = df_m['수량'].loc[line]
        df_input['단가4'].iloc[sumline]        = df_m['단가'].loc[line]
        df_input['공급가액4'].iloc[sumline]    = df_m['공급가액'].loc[line]
        df_input['세액4'].iloc[sumline]        = df_m['부가세'].loc[line]
        df_input['공급가액'].iloc[sumline]      = df_input['공급가액'].iloc[sumline] + df_m['공급가액'].loc[line]
        df_input['세액'].iloc[sumline]      = df_input['세액'].iloc[sumline] + df_m['부가세'].loc[line]


# ### 데이터 분석 
# * drop 기능으로 셀을 제거해서 사용하는 방법
# * 셀을 하나씩 증가시키며 데이터를 입력하는 방법

# In[10]:


#df_m1 = df_m.drop(0)
#df_m2 = df_m1.drop(1)
tax_name_old = df_m['전표번호'].iloc[0]
taxNewDataLineCount = 0
taxReleaseDataLineCount = 0

for i in range(df_tax_total):
    if(i == 0):
        TAX_NewDataInput(i, taxNewDataLineCount)
        TAX_ReleaseDataInput(i, taxReleaseDataLineCount, taxNewDataLineCount)
        taxNewDataLineCount = taxNewDataLineCount + 1
    else:
        if(tax_name_old != df_m['전표번호'].iloc[i]):       # 전표번호 변경 시 
            tax_name_old = df_m['전표번호'].iloc[i]         # 새로운 전표 등록
            
            taxReleaseDataLineCount = 0
            # 전표 신규 데이터 작성
            TAX_NewDataInput(i, taxNewDataLineCount)
            TAX_ReleaseDataInput(i, taxReleaseDataLineCount, taxNewDataLineCount)
            taxNewDataLineCount = taxNewDataLineCount + 1
        else:
            # 기존 전표번호에 데이터 추가
            taxReleaseDataLineCount = taxReleaseDataLineCount + 1
            TAX_ReleaseDataInput(i, taxReleaseDataLineCount, (taxNewDataLineCount - 1))
            #print("전표번호 같음", 'i =', i, '몇번째 품목 =', taxReleaseDataLineCount, '몇번째 라인 =', taxNewDataLineCount)


# ### 공급받는자 사업자 없으면 행 삭제

# In[11]:


#사업자 등록증 없는 업체는 삭제
#df_input.dropna(subset=['공급받는자 등록번호\n("-" 없이 입력)'], inplace=True)
df_input.dropna(subset=['세액1'], inplace=True)


# 그리고 만일 리스트 조건 안에 포함되는 데이터를 추출하고 싶다면 isin()함수를 사용해주면 된다. 만일, country가 [한국, 일본, 대만, 영국, 호주] 리스트에 포함되는 것을 추출하고 싶다면 기존에 조건을 여러개 열거했던 것처럼 사용하지 않고 아래 코드와 같은 형태를 사용함으로써 데이터를 추출할 수 있다.
# 
# country_list = ['한국', '일본', '대만', '영국', '호주']
# df[df['country'].isin(country_list)]
# 
# 반대로 country 열에 country_list에 포함되지 않는 데이터를 추출하고 싶다면 아래 코드 처럼 위의 코드에서 ~를 붙여 사용하면 된다.
# 
# country_list = ['한국', '일본', '대만', '영국', '호주']
# df[~df['country'].isin(country_list)]

# ### 계산서 별도 정리 업체 구분
# * 별도 업체는 사업자 등록번호로 조회 후 구분 할 것

# In[12]:


# unCompany_list = ['주식회사 쉰들러엘리베이터', '대한승강에이전트', '지에스이', '위드이엘', '신화엘리베이터㈜', '동양텍/울산']

# df_unCompany = df_input[~df_input['공급받는자 상호'].isin(unCompany_list)]
# df_Company = df_input[df_input['공급받는자 상호'].isin(unCompany_list)]

# #df_rs = df_input[df_input['공급받는자 상호'].str.contains('쉰들러')]
# #df_Company.to_excel('쉰들러 테스트.xlsx')

# #1. 파일 생성
# writer=pd.ExcelWriter('buffer.xlsx', engine='openpyxl')
 
# #2. 생성 파일에 시트명 지정 후 dataframe에 저장한 결과값 넣기
# df_Company.to_excel(writer, sheet_name='계산서 검토 업체')
# df_unCompany.to_excel(writer, sheet_name='계산서 발행 업체')
 
# #3. 작성 완료 후 파일 저장
# writer.save()


# In[13]:


'''
# 20221015 - 기존 파일 저장 방식
# df_input.to_excel("buffer.xlsx")
# wb = xl.load_workbook('buffer.xlsx')
# ws = wb.active
# #ws.insert_rows(0, 5)
# ws.delete_cols(1)
# wb.save('Tax List Complete File.xlsx')
# os.remove('buffer.xlsx')
# print("변환이 완료되었습니다.\n")
# #os.system("pause")
'''


# In[14]:


unCompany_list = ['주식회사 쉰들러엘리베이터', '대한승강에이전트', '지에스이', '위드이엘', '신화엘리베이터㈜', '동양텍/울산']

df_unCompany = df_input[~df_input['공급받는자 상호'].isin(unCompany_list)]
df_Company = df_input[df_input['공급받는자 상호'].isin(unCompany_list)]

#1. 파일 생성
writer=pd.ExcelWriter('Tax List Complete File.xlsx', engine='openpyxl')

#2. 생성 파일에 시트명 지정 후 dataframe에 저장한 결과값 넣기
df_Company.to_excel(writer, sheet_name='계산서 검토 업체', index = False)
df_unCompany.to_excel(writer, sheet_name='계산서 발행 업체', index = False)
#3. 작성 완료 후 파일 저장
writer.save()

print("변환이 완료되었습니다.\n")
#os.system("pause")


# # Jupyter File -> Python File Convertion
# jupyter nbconvert --to script Python_Excel_Conv_0V9000001.ipynb
# 
# # EXE 파일 생성
# pyinstaller -F .\Python_Excel_Conv_0V9000001.py
