import pandas as pd
from pandas import Series, DataFrame
from openpyxl.workbook import Workbook
import os, re, datetime
import sales_coupang as cp
import sales_naver as na
import sales_wmprice as wmp
import sales_gmarket as gm
import sales_auction as auc
import sales_11st as elev
import sales_tmon as tm
#import code_generator as cgen
import code_generator_new as cgen


def concat_all(path):
    # flist = [plist_f, qlist_f, nlist_f, namelist_f]
#    final_list = list(zip(cp.cal_coupang(path), na.cal_naver(path), wmp.cal_wmprice(path), gm.cal_gmarket(path)))

    local_dir = os.getcwd()
#    netdir = r'\\공용PC\메인컴퓨터\품절'
    netdir = 'C:/Python37/cherrycoco/excel_data/'
    os.chdir(netdir)

    final_cp = cp.cal_coupang(path)
#    print("end of cp")
    final_na = na.cal_naver(path)
#    print("end of na")
    final_wmp = wmp.cal_wmprice(path)
#    print("end of wmp")
    final_gm = gm.cal_gmarket(path)
#    print("end of gm")
    final_auc = auc.cal_auction(path)
#    print("end of auc")
    final_11st = elev.cal_11st(path)
#    print("end of 11st")
    final_tm = tm.cal_tmon(path)
#    print("end of tm")

    # 다시 로컬PC로 이동
    os.chdir(local_dir)


    fplist = final_cp[0] + final_na[0] + final_wmp[0] + final_gm[0] + final_auc[0] + final_11st[0] + final_tm[0]
    fqlist = final_cp[1] + final_na[1] + final_wmp[1] + final_gm[1] + final_auc[1] + final_11st[1] + final_tm[1]
    fnlist = final_cp[2] + final_na[2] + final_wmp[2] + final_gm[2] + final_auc[2] + final_11st[2] + final_tm[2]
    fnamelist = final_cp[3] + final_na[3] + final_wmp[3] + final_gm[3] + final_auc[3] + final_11st[3] + final_tm[3]

    # 최종 결과값 데이터프레임 생성
    df_all = pd.DataFrame({'상품코드': fplist, '상품명': fnamelist, '등록개수': fqlist, '주문수량': fnlist})
    df_all['등록개수'] = df_all['등록개수'].str.replace(pat='[개]|[xX]|[캔]|[매]|[(]|[)]', repl=r'', regex=True)

    df_final = cgen.code_generator()
    codes = list(df_final['품번'])
    product_name = list(df_final['제품명'])
    inbox = list(df_final['입수'])

    add_code, add_sold, add_name, add_box = [], [], [], []

    for pcode, pname, pbox in zip(codes, product_name, inbox):
#    print(pcode, pname)
        i = 0
        qsum = 0
        try:
            df_res_all = df_all[df_all['상품코드'] == pcode]
            res_plist = list(df_res_all['상품명'])
            for pn, n, m in zip(df_res_all['상품명'], df_res_all['등록개수'], df_res_all['주문수량']):
                i += 1
                if ',' not in n:
                    if len(set(res_plist)) > 1:
    #                if '-' not in res_plist and len(set(res_plist)) > 1:
    #                    if ',' in n:
    #                        idx = n.find(',')
    #                        n = n[(idx+1):]
                        if pn == pname:
                            qsum = qsum + int(n) * m
                        else: pass

                        if qsum > 0 and i == len(res_plist):
                            add_code.append(pcode)
                            add_sold.append(qsum)
                            add_name.append(pname)
                            add_box.append(int(pbox))
    #                        print(" *", pcode, pname, n, m, "=>", qsum)
                    else:
                        qsum = qsum + int(n) * m
                else: continue

            if len(set(res_plist)) == 1:
                if qsum > 0 and pn == pname:
                    add_code.append(pcode)
                    add_sold.append(qsum)
                    add_name.append(pn)
                    add_box.append(int(pbox))
    #                print(" **", pcode, pn, n, m, "=>", qsum)
    #        print("")
        except Exception as e:
            print("[ERROR_in_concat] " + pcode + ":" + str(e))
            continue


    # 최종 결과값 데이터프레임
    df_res = pd.DataFrame({'상품코드': [], '상품명': [], '입수': [], 'BOX': [], 'EA': []})
    df_res['상품코드'] = add_code
    df_res['상품명'] = add_name
    df_res['입수'] = add_box
    df_res['BOX'] = [x//y for x, y in zip(add_sold, add_box)]       # 몫연산
    df_res['EA'] = add_sold
    df_res = df_res.sort_values('상품코드')                   # 품번 기준 오름차순
#    df_res = df_res.sort_values('EA', ascending=False)  # 판매량 내림차순

    # 엑셀 파일로 저장
    filename = os.path.basename(path)
    filename = re.sub('[.][x]\w+', '', filename)  # 확장자 제거
    am_pm = filename[-4:-2]
    local_dir = os.getcwd()
    user_path = os.path.expanduser('~')
    save_dir = local_dir
#    save_dir = r'\\공용PC\메인컴퓨터\물류팀\입고파트\발주 재고리스트\판매량'
    os.chdir(save_dir)
    now = datetime.datetime.now()       # 오늘 날짜와 현재 시간
    today = str(now.date())             # 오늘 날짜(YYYY-MM-DD)
    today = re.sub('[-]', '', today[5:])
#    realtime = str(now.hour) + str(now.minute)      # 현재 시간
    df_res.to_excel('판매량_' + today + '_' + am_pm + '.xlsx', sheet_name='Sheet1', index=False)

    # 다시 로컬PC로 이동
    os.chdir(local_dir)
