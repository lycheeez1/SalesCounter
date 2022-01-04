import pandas as pd
from pandas import Series, DataFrame
import re, os, operator
#import code_generator as cgen
import code_generator_new as cgen
from openpyxl.workbook import Workbook
from difflib import SequenceMatcher

def match_rate(a, b):
    mrate = f'{SequenceMatcher(None, a, b).ratio()*100:.1f}%'
    mrate_n = float(re.sub('%', '', mrate))
    return mrate_n


def cal_coupang(path):
    df = pd.read_excel(path)
    df = df.fillna('-')

    # 유효한 열을 가공해서 새로운 열을 데이터프레임에 추가
    df['상품명N'] = df['판매자상품코드'].str.replace(pat='[+][발][롱]\w+|\d+[g]|[1][0][T]', repl=r'', regex=True)
    df['상품명N'] = df['상품명N'].str.replace(pat='[+][유][한][젠][(]\d+\w+\d+[)]|[(][임][의[증][정][)]', repl=r'', regex=True)
    df['상품명N'] = df['상품명N'].str.replace(pat='[+][코][코][사][은]\w+|[+][리][필]\d[.]\d[L]', repl=r'', regex=True)
    df['상품명N'] = df['상품명N'].str.replace(pat='\d+[개][입]', repl=r'', regex=True)
    df['상품코드'] = df['상품명']
    df['옵션N'] = df['주문선택사항'].str.replace(pat='[+][수][아]\w+[크]', repl=r'', regex=True)
#    df.drop(columns=['결제일','상태', '판매사이트 상품코드', '주문선택사항금액', '판매가', '배송비금액', '배송방법(원본)', '송장번호', '수집일'], inplace=True)


    ###########################################
    # code_generator로 품번 데이터프레임 생성 #
    ###########################################
    df_final = cgen.code_generator()
    code = list(df_final['품번'])
    product_name = list(df_final['제품명'])
    code_table = {}
    tmp = []
    # pc가 중복될 수 있음
    for pc, pn in zip(code, product_name):
        # pc값이 중복이면
        if pc in tmp:
            code_table[pc] += ', ' + pn
        else:
            code_table[pc] = pn
        tmp.append(pc)


    ################################################
    # 상품코드(p), 등록개수(q), 주문수량(n) 구하기 #
    ################################################
    platform = df['판매사이트명']
    df_cp = df[platform.str.contains('쿠팡|직접입력')]
#    df_cp = df[platform == '쿠팡(신)']

    plist, qlist, nlist, namelist = [], [], [], []
    plist_mul, qlist_mul,nlist_mul, oplist_mul = [], [], [], []

    for pn, pc, op, n in zip(df_cp['상품명N'], df_cp['상품코드'], df_cp['옵션N'], df_cp['주문수량']):
        # 상품코드
        product_code = re.compile('[a-z]{2}[0-9]{3}').findall(pc)      # 상품코드 (리스트)

        # 등록 개수
        try:
            if pn != '-':
                if product_code[0][0] == 'd':
                    pn = re.sub('\d+[m][l]|\d[.]\d[k][g]', '', pn)
                    register_quantity = re.compile('\d+').findall(pn)
                else:
                    if '개' in pn or 'x' in pn or '캔' in pn:
                        register_quantity = re.compile('\d+[개]|[xX]\d+|\d+[캔]').findall(pn)   # 등록 개수 (리스트)
                    else:
                        register_quantity = re.compile('\d+[매]').findall(pn)
                x = pn.count('+')
                y = pn.count(',')
                option = re.sub('\d+[개]|[xX]\d+|\d+[캔]|\d+[매]', '', pn)    # 나중에 품번 비교용
            else:
                if product_code[0][0] == 'd':
                    op = re.sub('\d+[m][l]|\d[.]\d[k][g]', '', op)
                    register_quantity = re.compile('\d+').findall(op)
                else:
                    register_quantity = re.compile('\d+[개]|[xX]\d+|\d+[캔]').findall(op)   # 등록 개수 (리스트)
                x = pn.count('+')
                y = pn.count(',')
                option = re.sub('\d+[개]|[xX]\d+|\d+[캔]|\d+[매]', '', op)    # 나중에 품번 비교용
        except IndexError as e1:
            print("[ERROR_00_cp_t1]", pc, "is out of range")
            continue
        except KeyError as e2:
            print("[ERROR_00_cp_t1]", str(e2), "is not available")
            continue

        if '+스너글' in option and '방향제' not in option:
            option = option.replace('+스너글', '+스너글 방향제')

        pcode = ",".join(product_code)           # 상품코드 (리스트->문자열)
        pcode = re.sub(' ', '', pcode)
        rquantity = ",".join(register_quantity)  # 등록개수 (리스트->문자열)

        # 단일
        try :
            if len(pcode) == 5:
                if rquantity == '':
                    rquantity = '1'

                rate = []
                if pcode in code_table.keys():
                    value = code_table[pcode]
                    if ',' in value:
                        vlist = value.split(',')
                        for v in vlist:
                            mrate = match_rate(v, option)
                            data = (pcode, mrate, v, rquantity, n)
                            rate.append(data)
                        ratelist = sorted(rate, key=operator.itemgetter(1), reverse=True)
                        rmax = ratelist[0]         # 문자열 일치율이 가장 높은 것
                        plist.append(rmax[0])
                        namelist.append(rmax[2].strip())
                        qlist.append(rmax[3])
                        nlist.append(rmax[4])
                    else:
                        plist.append(pcode)
                        namelist.append(value.strip())
                        qlist.append(rquantity)
                        nlist.append(n)
                else:
                    print("[ERROR_00_cp_if]", pcode, "is not available")
                    continue
            # 종합
            else:
                if rquantity == '':
                    if x > 0: rquantity = '1,' * (x+1)
                    elif y > 0: rquantity = '1,' * (y+1)
                    else: rquantity = '1'
                else:
                    if len(register_quantity) != (x+1):
                        if x > 0: rquantity = rquantity + ',1'
                    elif len(register_quantity) != (y+1):
                        if y > 0: rquantity = rquantity + ',1'
                ## 2, 1개 1, 2개 이런 거 보완
                #######

                rquantity = re.sub(r'[,]$', '', rquantity)  # 정규식으로

                plist_mul.append(pcode)
                qlist_mul.append(rquantity)
                nlist_mul.append(n)
                oplist_mul.append(option)

        except IndexError as e1:
            print("[ERROR_00_cp_t2]", pc, "is out of range")
            continue
        except KeyError as e2:
            print("[ERROR_00_cp_t2]", str(e2), "is not available")
            continue


    ##################################
    # 종합 주문 건 낱개로 풀어헤치기 #
    ##################################
    plist_2, qlist_2, nlist_2, namelist_2 = [], [], [], []
    plist_f, qlist_f, nlist_f, namelist_f = [], [], [], []

    # 종합 주문에서 정확히 뭘 골랐는지 찾아야 함 (상품코드 비교)
    for p, q, op, n in zip(plist_mul, qlist_mul, oplist_mul, nlist_mul):
        pl = p.split(",")           # 상품코드 리스트
        ql = q.split(",")           # 등록개수 리스트
        if '[' in op and '+' in op:
            op = op.replace('[', '{')
            op = op.replace(']', '}')
        if '르샤트라' in op and '(' in op and '+' in op:
            op = op.replace('(', '{')
            op = op.replace(')', '}')
        op = op.replace('++++', '')
        op = op.replace('+++', '')
        if '도브' in op and '워시' in op and '뷰티' in op and '너리싱' not in op:
            op = op.replace('뷰티', '뷰티 너리싱')
        if '몬스터' in op and '망고' in op and '로코' not in op:
            op = op.replace('망고', '망고로코')
        if '핫식스' in op and '슬릭' in op:
            op = re.sub('슬릭', '', op)
        if '럭스' in op and '로즈' in op:
            op = op.replace('로즈', '소프트핑크')
        if '립톤제로' in op:
            op = op.replace('립톤제로', '리퀴드 레몬&라임')
#        if '핫식스' in op and '슬릭' not in op:
#            op = op + '{뚱캔}'
#        if '핫식스' in op and '슬릭' in op:
#            op = re.sub('슬릭', '', op)
#            op = op + '{슬릭}'
        op = re.sub('[}][+]', ',', op)
        if '{' in op and '+' in op:
            idx = op.find('{')
            idx2 = op.find('+')
            pname = op[:idx]
            opl = re.split('[+]|[,]', op)
            if idx < idx2:
                for i, opp in enumerate(opl):
                    if pname not in opp and '스너글' not in opp:
                        opp_new = pname + opp
                    else:
                        opp_new = opp
                    opp_new = re.sub('[{]|[}]', '', opp_new)
                    opl.insert(i, opp_new)
                    opl.remove(opp)
        else:
    #        opl = op.split("+")
            opl = re.split('[+]|[,]', op)


        # 옵션 개수랑 등록 개수 맞추기
        if len(opl) < len(ql):
            while len(opl) < len(ql):
                ql = ql[:-1]

        global rate_list
        try:
            for pp in pl:
                if pp in code_table.keys():
                    pass
                else:
                    print("[ERROR_01_cp_for]", pp, "is not available")
                    pl.remove(pp)
            if pl[0] in code_table.keys():
                val = code_table[pl[0]]
                # 단일 옵션이고, 쌍 개수가 같고, 품번이 서로 다르고, 등록 개수가 모두 1이라면
                if ',' not in val and len(pl) == len(ql) and len(set(pl)) != 1 and len(set(ql)) == 1:
                    for pp, qq in zip(pl, ql):
                        val2 = code_table[pp]
                        plist_2.append(pp)
                        namelist_2.append(val2.strip())
                        qlist_2.append(qq)
                        nlist_2.append(n)
#                        print("if=>", pp, val2.strip(), qq, n)
                else:
                    rate_list = []
                    for pp in pl:
                        for opp, qq in zip(opl, ql):
                            if pp in code_table.keys():
                                value = code_table[pp]
                                # 한 품번에 옵션이 여러개인 경우
                                if ',' in value:
                                    vlist = value.split(',')
                                    for v in vlist:
                                        mrate = match_rate(v, opp)
#                                        print(mrate, "% | ", v, "| ", opp)
                                        data = (pp, mrate, v, qq, n)
                                        rate_list.append(data)
                                # 품번-상품명이 일대일대응인 경우
                                else:
                                    mrate = match_rate(value, opp)   # 딕셔너리 value값과 옵션의 문자열 일치
#                                    print(mrate, "% || ", value, "|| ", opp)
                                    data = (pp, mrate, value, qq, n)
                                    rate_list.append(data)

                                ratelist = sorted(rate_list, key=operator.itemgetter(1), reverse=True)

                                # 옵션 여러개인 놈 중복 제거
                                if ',' in value and len(set(pl)) == 1 and len(set(opl)) != 1:
                                    if ratelist[0][1] == ratelist[1][1]:
                                        ratelist = set(ratelist)
                                        ratelist = list(ratelist)
                                        ratelist.sort(reverse=True)
                            else:
                                print("[ERROR_01_cp_if]", pp, "is not available")
                                continue
                    try :
                        rlist = ratelist[:len(opl)]  # 내림차순한 일치율 리스트를 옵션 개수만큼 자름
    #                    print(rlist)
                        for r in rlist:
                            r = list(r)       # (pp, mrate, value, qq, n)
                            plist_2.append(r[0])
                            namelist_2.append(r[2].strip())
                            qlist_2.append(r[3])
                            nlist_2.append(r[4])
#                            print("else=>", r[0], r[2].strip(), r[3], r[4])
                    except Exception as e:
                        print("[ERROR_01_cp_t2]", repr(e))
#                        print("=>", pl, opl, ql)
                        continue
            else:
#                print("[ERROR_03_cp_if]", pl[0], "is not available")
                try:
                    pl = pl[1:]
                    for pp, qq in zip(pl, ql):
                        val2 = code_table[pp]
                        plist_2.append(pp)
                        namelist_2.append(val2.strip())
                        qlist_2.append(qq)
                        nlist_2.append(n)
#                        print("if2=>", pp, val2.strip(), qq, n)
                except Exception as e:
                    print("[ERROR_01_cp_t3]", repr(e))
                    continue
        # 품번 없을 시 예외 (KeyError로 인한 Runtime에러 방지)
        except Exception as e:
            print("[ERROR_01_cp_t1]", repr(e))
#            print("=>", pl, opl, ql)
#            rate_list.clear()
            continue

    plist_f = plist + plist_2
    qlist_f = qlist + qlist_2
    nlist_f = nlist + nlist_2
    namelist_f = namelist + namelist_2

    flist = [plist_f, qlist_f, nlist_f, namelist_f]
    return flist

'''
    # 새로운 데이터프레임 생성
    df_coupang = pd.DataFrame({'상품코드': plist_f, '상품명': namelist_f, '등록개수': qlist_f, '주문수량': nlist_f})
    df_coupang['등록개수'] = df_coupang['등록개수'].str.replace(pat='[개]|[xX]|[캔]', repl=r'', regex=True)

    df_sales_cp = cdf.create_df(df_coupang, df_final)      # 쿠팡 판매량 데이터프레임

    return df_sales_cp
'''

    #####
