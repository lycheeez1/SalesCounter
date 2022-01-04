import pandas as pd
from pandas import Series, DataFrame
#import datetime as dt
import os


def code_generator():
    local_dir = os.getcwd()

#    netdir = r'\\공용PC\Users\Documents\물류팀\입고파트\발주 재고리스트'
    netdir = 'C:/Python37/cherrycoco/excel_data/Gmail/'
    os.chdir(netdir)

    # --- 생활 & 뷰티 일반 --- * -
    smbt = pd.ExcelFile('2021 SM화장품/2021 SM화장품(뷰티).xlsx')
    sheets_smbt = smbt.sheet_names
    if 'Sheet' in sheets_smbt[-1]:
        last_sheet = sheets_smbt[-3]
    else:
        last_sheet = sheets_smbt[-2]
    df_smbt = pd.read_excel('2021 SM화장품/2021 SM화장품(뷰티).xlsx', sheet_name=last_sheet, usecols=[0, 2, 4], skiprows=3)
    df_smbt.columns = ['품번', '제품명', '입수']

    smlf = pd.ExcelFile('2021 SM화장품/2021 SM화장품(생활).xlsx')
    sheets_smlf = smlf.sheet_names
    if 'Sheet' in sheets_smlf[-1]:
        last_sheet = sheets_smlf[-4]
    else:
        last_sheet = sheets_smlf[-3]
    df_smlf = pd.read_excel('2021 SM화장품/2021 SM화장품(생활).xlsx', sheet_name=last_sheet, usecols=[0, 1, 4], skiprows=3)
    df_smlf.columns = ['품번', '제품명', '입수']

    sunghwa = pd.ExcelFile("2021 성화유통/2021 성화 재고리스트.xlsx")
    sheets_sunghwa = sunghwa.sheet_names
    if 'Sheet' in sheets_sunghwa[-1]:
        last_sheet = sheets_sunghwa[-4]
    else:
        last_sheet = sheets_sunghwa[-3]
    df_sunghwa = pd.read_excel("2021 성화유통/2021 성화 재고리스트.xlsx", sheet_name=last_sheet, usecols=[0, 1, 2], skiprows=3)
    df_sunghwa.columns = ['품번', '제품명', '입수']


    # --- 로드샵 --- *-
    nature = pd.ExcelFile("2021 네이쳐리퍼블릭 재고리스트.xlsx")
    sheets_nature = nature.sheet_names
    if 'Sheet' in sheets_nature[-1]:
        last_sheet = sheets_nature[-2]
    else:
        last_sheet = sheets_nature[-1]
    df_nature = pd.read_excel("2021 네이쳐리퍼블릭 재고리스트.xlsx", sheet_name=last_sheet, usecols=[0, 1, 10], skiprows=4)
    df_nature.columns = ['품번', '제품명', '입수']

    saem = pd.ExcelFile("2021 더샘 재고리스트.xlsx")
    sheets_saem = saem.sheet_names
    if 'Sheet' in sheets_saem[-1]:
        last_sheet = sheets_saem[-2]
    else:
        last_sheet = sheets_saem[-1]
    df_saem = pd.read_excel("2021 더샘 재고리스트.xlsx", sheet_name=last_sheet, usecols=[0, 1, 10], skiprows=4)
    df_saem.columns = ['품번', '제품명', '입수']

    fshop = pd.ExcelFile("2021 더페이스샵 재고리스트.xlsx")
    sheets_fshop = fshop.sheet_names
    if 'Sheet' in sheets_fshop[-1]:
        last_sheet = sheets_fshop[-2]
    else:
        last_sheet = sheets_fshop[-1]
    df_fshop = pd.read_excel("2021 더페이스샵 재고리스트.xlsx", sheet_name=last_sheet, usecols=[0, 1, 8], skiprows=4)
    df_fshop.columns = ['품번', '제품명', '입수']

    skinfood = pd.ExcelFile("2021 스킨푸드 재고리스트.xlsx")
    sheets_skinfood = skinfood.sheet_names
    if 'Sheet' in sheets_skinfood[-1]:
        last_sheet = sheets_skinfood[-2]
    else:
        last_sheet = sheets_skinfood[-1]
    df_skinfood = pd.read_excel("2021 스킨푸드 재고리스트.xlsx", sheet_name=last_sheet, usecols=[0, 1, 6], skiprows=4)
    df_skinfood.columns = ['품번', '제품명', '입수']

    etude = pd.ExcelFile("2021 에뛰드 재고리스트.xlsx")
    sheets_etude = etude.sheet_names
    if 'Sheet' in sheets_etude[-1]:
        last_sheet = sheets_etude[-2]
    else:
        last_sheet = sheets_etude[-1]
    df_etude = pd.read_excel("2021 에뛰드 재고리스트.xlsx", sheet_name=last_sheet, usecols=[0, 1, 3], skiprows=4)
    df_etude.columns = ['품번', '제품명', '입수']

    inni = pd.ExcelFile("2021 이니스프리 재고리스트.xlsx")
    sheets_inni = inni.sheet_names
    if 'Sheet' in sheets_inni[-1]:
        last_sheet = sheets_inni[-2]
    else:
        last_sheet = sheets_inni[-1]
    df_inni = pd.read_excel("2021 이니스프리 재고리스트.xlsx", sheet_name=last_sheet, usecols=[0, 1, 5], skiprows=4)
    df_inni.columns = ['품번', '제품명', '입수']

    tony = pd.ExcelFile("2021 토니모리 재고리스트.xlsx")
    sheets_tony = tony.sheet_names
    if 'Sheet' in sheets_tony[-1]:
        last_sheet = sheets_tony[-2]
    else:
        last_sheet = sheets_tony[-1]
    df_tony = pd.read_excel("2021 토니모리 재고리스트.xlsx", sheet_name=last_sheet, usecols=[0, 1, 6], skiprows=4)
    df_tony.columns = ['품번', '제품명', '입수']

    holika = pd.ExcelFile("2021 홀리카홀리카 재고리스트.xlsx")
    sheets_holika = holika.sheet_names
    if 'Sheet' in sheets_holika[-1]:
        last_sheet = sheets_holika[-2]
    else:
        last_sheet = sheets_holika[-1]
    df_holika = pd.read_excel("2021 홀리카홀리카 재고리스트.xlsx", sheet_name=last_sheet, usecols=[0, 1, 7], skiprows=4)
    df_holika.columns = ['품번', '제품명', '입수']


    # --- 음료 --- * -
    lotte = pd.ExcelFile("2021 롯데칠성/2021 롯데칠성 ★재고리스트.xlsx")
    sheets_lotte = lotte.sheet_names
    if 'Sheet' in sheets_lotte[-1]:
        last_sheet = sheets_lotte[-3]
    else:
        last_sheet = sheets_lotte[-2]
    df_lotte = pd.read_excel("2021 롯데칠성/2021 롯데칠성 ★재고리스트.xlsx", sheet_name=last_sheet, usecols=[0, 1, 4], skiprows=3)
    df_lotte.columns = ['품번', '제품명', '입수']

    cola = pd.ExcelFile("2021 코카콜라/2021 코카콜라 ★재고리스트.xlsx")
    sheets_cola = cola.sheet_names
    if 'Sheet' in sheets_cola[-1]:
        last_sheet = sheets_cola[-3]
    else:
        last_sheet = sheets_cola[-2]
    df_cola = pd.read_excel("2021 코카콜라/2021 코카콜라 ★재고리스트.xlsx", sheet_name=last_sheet, usecols=[0, 2, 5], skiprows=3)
    df_cola.columns = ['품번', '제품명', '입수']

    # --- 시세이도 --- * -
    shis = pd.ExcelFile(netdir +"2021 시세이도 재고리스트.xlsx")
    sheets_shis = shis.sheet_names
    if 'Sheet' in sheets_shis[-1]:
        last_sheet = sheets_shis[-3]
    else:
        last_sheet = sheets_shis[-2]
    df_shi = pd.read_excel(netdir +"2021 시세이도 재고리스트.xlsx", sheet_name=last_sheet, usecols=[0, 1], skiprows=3)
    list_sh = ['1'] * len(df_shi)
    ser_sh = pd.Series(data=list_sh)
    df_sh = ser_sh.to_frame()
    df_shis = df_shi.join(df_sh)
    df_shis.columns = ['품번','제품명','입수']

    # 다시 로컬PC로 이동
    os.chdir(local_dir)

    # --- 전체 데이터프레임 concat ---- * -
    df_final = pd.concat([df_smbt,df_smlf, df_sunghwa, df_nature, df_saem,
                          df_fshop, df_skinfood, df_etude, df_inni, df_tony, df_holika,
                          df_lotte, df_cola, df_shis], ignore_index=True)
    df_final.dropna(axis=0, inplace=True)
    #df = pd.concat([s_2016, s_2017, s_2018], keys=['2016', '2017', '2018'])
    # 위처럼 하면 브랜드별로 인덱스 달 수 있음

    # 엑셀 파일로 저장
#    save_dir = "C:/Users/bjh/Desktop/"
#    now = dt.datetime.now()
#    now_MMDD = now.strftime("%m%d")
#    df_final.to_excel(save_dir + '품번 테이블_' + now_MMDD + '.xlsx', sheet_name='Sheet1',index=False)

    return df_final
