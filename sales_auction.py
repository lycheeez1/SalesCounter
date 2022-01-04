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


def cal_auction(path):
    df = pd.read_excel(path)
    df = df.fillna('-')

    # 유효한 열을 가공해서 새로운 열을 데이터프레임에 추가
    df['상품명N'] = df['상품명'].str.replace(pat='[+][스][너][글]|[+][허][브][티]||[1][0][T]', repl=r'', regex=True)
    #df['상품명N'] = df['상품명N'].str.replace(pat='\d[.]\d[k][g]|[/]|[+][+][+]', repl=r'', regex=True)
    df['옵션N'] = df['주문선택사항'].str.replace(pat=r' ', repl=r'', regex=True)
    df['옵션N'] = df['옵션N'].str.replace(pat='[/]\d+[개원]', repl=r'', regex=True)
    df['옵션N'] = df['옵션N'].str.replace(pat='[선]\w+[:]|[색][상][:]|[발][송][일][:]', repl=r'', regex=True)
    df['옵션N'] = df['옵션N'].str.replace(pat='[종][류][:]|[라][인][:]|[사][은][품][:]', repl=r'', regex=True)
    #df['옵션N'] = df['옵션N'].str.replace(pat='[선][택]\d[:]|[선][택][:]', repl=r'', regex=True)
    df['상품코드'] = df['판매자상품코드']


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


    ##########################
    # 딜지 데이터프레임 가공 #
    ##########################
    excel_url = "종합 품절_오픈.xlsm"
    deal_smart = pd.read_excel(excel_url, sheet_name=None)
    klist = list(deal_smart.keys())[1:]
    vlist = list(deal_smart.values())[1:]
    vlist_new = []

    for k, v in zip(klist, vlist):
        v.insert(0, '딜번호', k)
        v = v.iloc[7:, 0:6]
        v.columns = ['딜번호', '옵션#','널1', '널2', '널3', '품번']
        v.drop(['널1', '널2', '널3'], axis=1, inplace=True)
        vlist_new.append(v)

    deal_dic = dict(zip(klist, vlist_new))

    concat_deal = pd.concat(deal_dic, ignore_index=True)
    concat_deal.dropna(axis=0, inplace=True)
    delete = concat_deal['옵션#'].isin(['#'])
    concat_deal = concat_deal[~delete]


    ################################################
    # 상품코드(p), 등록개수(q), 주문수량(n) 구하기 #
    ################################################
    platform = df['판매사이트명']
    df_auc = df[platform.str.contains('옥션')]

    plist, qlist, nlist, namelist = [], [], [], []
    plist_mul, qlist_mul, nlist_mul, oplist_mul = [], [], [], []

    for pn, pc, op, n in zip(df_auc['상품명N'], df_auc['상품코드'], df_auc['옵션N'], df_auc['주문수량']):
        # 상품코드
        try:
            product_code = []
            # 종합 딜
            if len(pc) == 4:
                deal_code = re.compile('[a-z]{2}[0-9]{2}').findall(pc)      # 딜코드 (리스트)
                option_num = re.compile('[0-9]{2}[)]').findall(op)          # 상품 옵션 번호 (리스트)

                op = re.sub('[가-힣][_][가-힣]', '', op)
                ix = op.count('_')
                if len(option_num) > (ix+1):
                    option_num = option_num[:(ix+1)]

                for i, opt in enumerate(option_num):
                    option_num[i] = opt[:-1]

                deal_df = concat_deal[concat_deal['딜번호'] == deal_code[0]]
                for opt in option_num:
                    res = deal_df[deal_df['옵션#'] == opt]
                    pdcode = res['품번']
                    product_code.extend(list(pdcode))
                    product_code = list(set(product_code))

            # 단일 & 코드 나열 종합
            else:
                product_code = re.compile('[a-z]{2}[0-9]{3}').findall(pc)
        except IndexError as e1:
            print("[ERROR_00_auc_t1]", pc, "is out of range")
            continue
        except KeyError as e2:
            print("[ERROR_00_auc_t1]", str(e2), "is not available")
            continue


        # 등록 개수
        try:
            if op == '-':
                if product_code[0][0] == 'd':
                    pn = re.sub('\d+[m][l]|\d[.]\d[k][g]', '', pn)
                    register_quantity = re.compile('\d+').findall(pn)
                else:
                    register_quantity = re.compile('\d+[개]|[xX]\d+|\d+[캔]').findall(pn)   # 등록 개수 (리스트)
                x = pn.count('+')
                y = pn.count('_')
                if 'SPF' in pn and '+' in pn:
                    x = x - 1
                option = re.sub('\d+[개]|[xX]\d+|\d+[캔]', '', pn)    # 나중에 품번 비교용
                option = re.sub('[(][)]', '', option)
            else:
                op = re.sub(r'[)]$', '', op)    # 옵션 마지막의 ) 없애기 (옵션넘버랑 안 헷갈리게)
                op = re.sub('[0-2][0-9][)]', '', op)       # 갯수 말고 옵션만 삭제하기 위해
                op = op.replace('}+{', '+')
                if product_code[0][0] == 'd':
                    op = re.sub('\d+[m][l]|\d[.]\d[k][g]', '', op)
                    register_quantity = re.compile('\d+').findall(op)
                else:
                    register_quantity = re.compile('\d+[개]|[xX]\d+|\d+[캔]').findall(op)   # 등록 개수 (리스트)
                x = op.count('+')
                y = op.count('_')
                option = re.sub('\d+[개]|[xX]\d+|\d+[캔]', '', op)    # 나중에 품번 비교용
                option = re.sub('[(][)]', '', option)

        except IndexError as e1:
            print("[ERROR_00_auc_t2]", pc, "is out of range")
            continue
        except KeyError as e2:
            print("[ERROR_00_auc_t2]", str(e2), "is not available")
            continue

        pcode = ",".join(product_code)           # 상품코드 (리스트->문자열)
        pcode = re.sub(' ', '', pcode)
        rquantity = ",".join(register_quantity)  # 등록개수 (리스트->문자열)


        # 단일 (or상품코드가 같은 종합)
        if len(pcode) == 5:
            if rquantity == '':
                if x == 0 and y == 0:
                    rquantity = '1'
                else:
                    if x > 0:
                        pcode = (pcode + ',') * (x+1)
                        rquantity = '1,' * (x+1)
                    elif y > 0:
                        pcode = (pcode + ',') * (y+1)
                        rquantity = '1,' * (y+1)
            else:
                if len(register_quantity) != (x+1):
                    if x > 0:
                        rquantity = rquantity + ',1'
                elif len(register_quantity) != (y+1):
                    if y > 0:
                        rquantity = rquantity + ',1'

                if rquantity.count(',') >= 0:
                    if x > 0: pcode = (pcode + ',') * (x+1)
                    if y > 0: pcode = (pcode + ',') * (y+1)

            pcode = re.sub(r'[,]$', '', pcode)  # 정규식으로
            rquantity = re.sub(r'[,]$', '', rquantity)  # 정규식으로
            try:
                if len(pcode) == 5:
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
                        print("[ERROR_00_auc_if]", pcode, "is not available")
                        continue
                else:
                    plist_mul.append(pcode)
                    qlist_mul.append(rquantity)
                    nlist_mul.append(n)
                    oplist_mul.append(option)
            except Exception as e:
                print("[ERROR_00_auc_t3]", str(e))
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


    ##################################
    # 종합 주문 건 낱개로 풀어헤치기 #
    ##################################
    plist_2, qlist_2, nlist_2, namelist_2 = [], [], [], []
    plist_f, qlist_f, nlist_f, namelist_f = [], [], [], []

    for p, q, op, n in zip(plist_mul, qlist_mul, oplist_mul, nlist_mul):
        pl = p.split(",")           # 상품코드 리스트
        ql = q.split(",")           # 등록개수 리스트
        op = op.replace('++++', '')
        op = op.replace('+++', '')
        op = re.sub('[}][+]', ',', op)
        # --- 치환 --- * -
        if '컨디' in op and '컨디셔너' not in op:
            op = op.replace('컨디', '컨디셔너')
        if '도브' in op and '워시' in op and '뷰티' in op and '너리싱' not in op:
            op = op.replace('뷰티', '뷰티 너리싱')
        if '몬스터' in op and '망고' in op and '로코' not in op:
            op = op.replace('망고', '망고로코')
        if '럭스' in op and '로즈' in op:
            op = op.replace('로즈', '소프트핑크')
        if '립톤제로' in op:
            op = op.replace('립톤제로', '리퀴드 레몬&라임')
        if '{' in op and '+' in op:
            idx = op.find('{')
            idx2 = op.find('+')
            pname = op[:idx]
            opl = re.split('[+]|[,]|[_]', op)
            if idx < idx2:
                for i, opp in enumerate(opl):
                    if pname not in opp:
                        opp_new = pname + opp
                    else:
                        opp_new = opp
                    opp_new = re.sub('[{]|[}]', '', opp_new)
                    opl.insert(i, opp_new)
                    opl.remove(opp)
        else:
            opl = re.split('[+]|[,]|[_]', op)


        # 옵션 개수랑 등록 개수 맞추기
        if len(opl) < len(ql):
            while len(opl) < len(ql):
                ql = ql[:-1]


        rate_list, rlist = [], []
        try:
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
                    for pp in pl:
                        for opp, qq in zip(opl, ql):
                            rate_list.clear()
                            if pp in code_table.keys():
                                value = code_table[pp]
                                # 한 품번에 옵션이 여러개인 경우
                                if ',' in value:
                                    vlist = value.split(',')
                                    for v in vlist:
                                        mrate = match_rate(v, opp)
#                                        print("if", mrate, "% | ", v, "| ", opp)
                                        data = (pp, mrate, v, qq, n)
                                        rate_list.append(data)
                                # 품번-상품명이 일대일대응인 경우
                                else:
                                    mrate = match_rate(value, opp)   # 딕셔너리 value값과 옵션의 문자열 일치
#                                    print("else", mrate, "% || ", value, "|| ", opp)
                                    data = (pp, mrate, value, qq, n)
                                    rate_list.append(data)

                                ratelist = sorted(rate_list, key=operator.itemgetter(1), reverse=True)
                                rlist.append(ratelist[0])
                            else:
                                print("[ERROR_02_auc_if]", pp, "is not available")
                                continue
                        if len(set(pl)) == 1:
                            break
                    try :
#                        rlist = ratelist[:len(opl)]  # 내림차순한 일치율 리스트를 옵션 개수만큼 자름
                        for r in rlist:
                            r = list(r)       # (pp, mrate, value, qq, n)
                            plist_2.append(r[0])
                            namelist_2.append(r[2].strip())
                            qlist_2.append(r[3])
                            nlist_2.append(r[4])
#                            print("else=>", r[0], r[2].strip(), r[3], r[4])
                    except Exception as e:
                        print("[ERROR_01_auc_t2]", str(e))
#                        print("=>", pl, opl, ql)
                        continue
            else:
                print("[ERROR_03_auc_if]", pl[0], "is not available")
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
                    print("[ERROR_01_auc_t3]", str(e))
                    continue
        # 품번 없을 시 예외 (KeyError로 인한 Runtime에러 방지)
        except Exception as e:
            print("[ERROR_01_t1]", str(e))
#            print("=>", pl, opl, ql)
#            rate_list.clear()
            continue

    plist_f = plist + plist_2
    qlist_f = qlist + qlist_2
    nlist_f = nlist + nlist_2
    namelist_f = namelist + namelist_2

    flist = [plist_f, qlist_f, nlist_f, namelist_f]
    return flist
