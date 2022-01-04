import pandas as pd
from pandas import Series, DataFrame
#import datetime as dt
import os


def code_generator():
    local_dir = os.getcwd()

#    netdir = r'\\공용PC\메인컴퓨터\물류팀\입고파트\발주 재고리스트'
    netdir = 'C:/Python37/cherrycoco/excel_data/Gmail/'
    os.chdir(netdir)

    # --- 생활 & 뷰티 일반 --- * -
    smbt = pd.ExcelFile('2021 SM화장품/2021 SM화장품(뷰티).xlsx')
    sheets_smbt = smbt.sheet_names
    if 'Sheet' in sheets_smbt[-1]:
        last_sheet = sheets_smbt[-3]
    else:
        last_sheet = sheets_smbt[-2]
    df_smbt = pd.read_excel('2021 SM화장품/2021 SM화장품(뷰티).xlsx', sheet_name=last_sheet)
    df_smbt = df_smbt.iloc[4:, 0:5]
    df_smbt.columns = ['품번', '브랜드', '제품명', '단위', '입수']
    df_smbt.drop(['브랜드', '단위'], axis=1, inplace=True)

    smlf = pd.ExcelFile('2021 SM화장품/2021 SM화장품(생활).xlsx')
    sheets_smlf = smlf.sheet_names
    if 'Sheet' in sheets_smlf[-1]:
        last_sheet = sheets_smlf[-4]
    else:
        last_sheet = sheets_smlf[-3]
    df_smlf = pd.read_excel('2021 SM화장품/2021 SM화장품(생활).xlsx', sheet_name=last_sheet)
    df_smlf = df_smlf.iloc[4:, 0:5]
    df_smlf.columns = ['품번','제품명', '널', '유통기한', '입수']
    df_smlf.drop(['널', '유통기한'], axis=1, inplace=True)

    sunghwa = pd.ExcelFile("2021 성화유통/2021 성화 재고리스트.xlsx")
    sheets_sunghwa = sunghwa.sheet_names
    if 'Sheet' in sheets_sunghwa[-1]:
        last_sheet = sheets_sunghwa[-4]
    else:
        last_sheet = sheets_sunghwa[-3]
    df_sunghwa = pd.read_excel("2021 성화유통/2021 성화 재고리스트.xlsx", sheet_name=last_sheet)
    df_sunghwa = df_sunghwa.iloc[4:, 0:3]
    df_sunghwa.columns = ['품번','제품명', '입수']
    df_sunghwa.loc[(df_sunghwa.품번 == 'cg004'), '입수'] = 12
    df_sunghwa.loc[(df_sunghwa.품번 == 'ca012'), '입수'] = 54
    df_sunghwa.loc[(df_sunghwa.품번 == 'ca013'), '입수'] = 54


    # --- 로드샵 --- *-
    nature = pd.ExcelFile("2021 네이쳐리퍼블릭 재고리스트.xlsx")
    sheets_nature = nature.sheet_names
    if 'Sheet' in sheets_nature[-1]:
        last_sheet = sheets_nature[-2]
    else:
        last_sheet = sheets_nature[-1]
    df_nature = pd.read_excel("2021 네이쳐리퍼블릭 재고리스트.xlsx", sheet_name=last_sheet)
    df_n1 = df_nature.iloc[5:, 0:2]
    df_n2 = df_nature.iloc[5:, 10]
    df_nat = df_n1.join(df_n2)
    df_nat.columns = ['품번','제품명', '입수']

    saem = pd.ExcelFile("2021 더샘 재고리스트.xlsx")
    sheets_saem = saem.sheet_names
    if 'Sheet' in sheets_saem[-1]:
        last_sheet = sheets_saem[-2]
    else:
        last_sheet = sheets_saem[-1]
    df_saem = pd.read_excel("2021 더샘 재고리스트.xlsx", sheet_name=last_sheet)
    df_sae1 = df_saem.iloc[5:, 0:2]
    df_sae2 = df_saem.iloc[5:, 4]
    df_saemm = df_sae1.join(df_sae2)
    df_saemm.columns = ['품번','제품명','입수']

    fshop = pd.ExcelFile("2021 더페이스샵 재고리스트.xlsx")
    sheets_fshop = fshop.sheet_names
    if 'Sheet' in sheets_fshop[-1]:
        last_sheet = sheets_fshop[-2]
    else:
        last_sheet = sheets_fshop[-1]
    df_fshop = pd.read_excel("2021 더페이스샵 재고리스트.xlsx", sheet_name=last_sheet)
    df_f1 = df_fshop.iloc[5:, 0:2]
    df_f2 = df_fshop.iloc[5:, 8]
    df_face = df_f1.join(df_f2)
    df_face.columns = ['품번', '제품명', '입수']

    skinfood = pd.ExcelFile("2021 스킨푸드 재고리스트.xlsx")
    sheets_skinfood = skinfood.sheet_names
    if 'Sheet' in sheets_skinfood[-1]:
        last_sheet = sheets_skinfood[-2]
    else:
        last_sheet = sheets_skinfood[-1]
    df_skinfood = pd.read_excel("2021 스킨푸드 재고리스트.xlsx", sheet_name=last_sheet)
    df_sk1 = df_skinfood.iloc[5:, 0:2]
    df_sk2 = df_skinfood.iloc[5:, 6]
    df_skin = df_sk1.join(df_sk2)
    df_skin.columns = ['품번','제품명','입수']

    etude = pd.ExcelFile("2021 에뛰드 재고리스트.xlsx")
    sheets_etude = etude.sheet_names
    if 'Sheet' in sheets_etude[-1]:
        last_sheet = sheets_etude[-2]
    else:
        last_sheet = sheets_etude[-1]
    df_etude = pd.read_excel("2021 에뛰드 재고리스트.xlsx", sheet_name=last_sheet)
    df_et1 = df_etude.iloc[5:, 0:2]
    df_et2 = df_etude.iloc[5:, 3]
    df_etu = df_et1.join(df_et2)
    df_etu.columns = ['품번','제품명','입수']

    inni = pd.ExcelFile("2021 이니스프리 재고리스트.xlsx")
    sheets_inni = inni.sheet_names
    if 'Sheet' in sheets_inni[-1]:
        last_sheet = sheets_inni[-2]
    else:
        last_sheet = sheets_inni[-1]
    df_inni = pd.read_excel("2021 이니스프리 재고리스트.xlsx", sheet_name=last_sheet)
    df_in1 = df_inni.iloc[5:, 0:2]
    df_in2 = df_inni.iloc[5:, 5]
    df_innis = df_in1.join(df_in2)
    df_innis.columns = ['품번','제품명','입수']

    tony = pd.ExcelFile("2021 토니모리 재고리스트.xlsx")
    sheets_tony = tony.sheet_names
    if 'Sheet' in sheets_tony[-1]:
        last_sheet = sheets_tony[-2]
    else:
        last_sheet = sheets_tony[-1]
    df_tony = pd.read_excel("2021 토니모리 재고리스트.xlsx", sheet_name=last_sheet)
    df_tn1 = df_tony.iloc[5:, 0:2]
    df_tn2 = df_tony.iloc[5:, 6]
    df_tnml = df_tn1.join(df_tn2)
    df_tnml.columns = ['품번','제품명','입수']

    holika = pd.ExcelFile("2021 홀리카홀리카 재고리스트.xlsx")
    sheets_holika = holika.sheet_names
    if 'Sheet' in sheets_holika[-1]:
        last_sheet = sheets_holika[-2]
    else:
        last_sheet = sheets_holika[-1]
    df_holika = pd.read_excel("2021 홀리카홀리카 재고리스트.xlsx", sheet_name=last_sheet)
    df_hl1 = df_holika.iloc[5:, 0:2]
    df_hl2 = df_holika.iloc[5:, 7]
    df_holi = df_hl1.join(df_hl2)
    df_holi.columns = ['품번','제품명','입수']


    # --- 음료 --- * -
    lotte = pd.ExcelFile("2021 롯데칠성/2021 롯데칠성 ★재고리스트.xlsx")
    sheets_lotte = lotte.sheet_names
    if 'Sheet' in sheets_lotte[-1]:
        last_sheet = sheets_lotte[-3]
    else:
        last_sheet = sheets_lotte[-2]
    df_lotte = pd.read_excel("2021 롯데칠성/2021 롯데칠성 ★재고리스트.xlsx", sheet_name=last_sheet)
    df_lt1 = df_lotte.iloc[4:, 0:2]
    df_lt2 = df_lotte.iloc[4:, 4]
    df_lott = df_lt1.join(df_lt2)
    df_lott.columns = ['품번','제품명','입수']

    cola = pd.ExcelFile("2021 코카콜라/2021 코카콜라 ★재고리스트.xlsx")
    sheets_cola = cola.sheet_names
    if 'Sheet' in sheets_cola[-1]:
        last_sheet = sheets_cola[-3]
    else:
        last_sheet = sheets_cola[-2]
    df_cola = pd.read_excel("2021 코카콜라/2021 코카콜라 ★재고리스트.xlsx", sheet_name=last_sheet)
    df_col1 = df_cola.iloc[4:, 0:3]
    df_col2 = df_cola.iloc[4:, 5]
    df_coca = df_col1.join(df_col2)
    df_coca.columns = ['품번', '브랜드', '제품명', '입수']
    df_coca.drop(['브랜드'], axis=1, inplace=True)


    # --- 시세이도 --- * -
    shis = pd.ExcelFile("2021 시세이도/2021 시세이도 재고리스트.xlsx")
    sheets_shis = shis.sheet_names
    if 'Sheet' in sheets_shis[-1]:
        last_sheet = sheets_shis[-3]
    else:
        last_sheet = sheets_shis[-2]
    df_shis = pd.read_excel("2021 시세이도/2021 시세이도 재고리스트.xlsx", sheet_name=last_sheet)
    df_sh1 = df_shis.iloc[4:, 0:2]
    list_sh2 = ['1'] * len(df_sh1)
    ser_sh2 = pd.Series(data=list_sh2)
    df_sh2 = ser_sh2.to_frame()
    df_shi = df_sh1.join(df_sh2)
    df_shi.columns = ['품번','제품명','입수']


    # --- 식품 --- * -
    ottugi = pd.ExcelFile("2021 오뚜기/2021 오뚜기 ★재고리스트.xlsx")
    sheets_ott = ottugi.sheet_names
    if 'Sheet' in sheets_ott[-1]:
        last_sheet = sheets_ott[-2]
    else:
        last_sheet = sheets_ott[-1]
    df_ottugi = pd.read_excel("2021 오뚜기/2021 오뚜기 ★재고리스트.xlsx", sheet_name=last_sheet)
    df_ott1 = df_ottugi.iloc[5:, 0:3]
    df_ott2 = df_ottugi.iloc[5:, 5]
    df_ott = df_ott1.join(df_ott2)
    df_ott.columns = ['품번', '브랜드', '제품명', '입수']
    df_ott.drop(['브랜드'], axis=1, inplace=True)


    # 다시 로컬PC로 이동
    os.chdir(local_dir)

    # --- 전체 데이터프레임 concat ---- * -
    df_final = pd.concat([df_smbt, df_smlf, df_sunghwa, df_nat, df_saemm,
                          df_face, df_skin, df_etu, df_innis, df_tnml, df_holi,
                          df_lott, df_coca, df_shi, df_ott], ignore_index=True)
    df_final.dropna(axis=0, inplace=True)
    delete = df_final['품번'].isin(['품번'])
    df_final = df_final[~delete]
#    df_final.astype({'입수':'int'})

    return df_final
