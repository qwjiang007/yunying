import pandas as pd
import os
import xlwings as xw


this_file = os.path.realpath(__file__)
#print(this_file)
#this_file = 'E:\PythonCode\运营出入金/新股缴款/'
#存放原始文件的文件夹
source = os.path.join(os.path.dirname(this_file), '原始文件')
#存放解析好的文件
output = os.path.join(os.path.dirname(this_file), '输出文件')


for filename in os.listdir(source):
    if "~$" in filename:
        continue
    #打开需要解析的文件
    org_info = pd.read_excel(os.path.join(source, filename), skiprows=2)
    cus_lst = list(set(org_info['基金托管人'].tolist()))
    print('托管',cus_lst)
    for cus in cus_lst:
        print('托管', cus)
        sub_info = org_info[org_info['基金托管人'] == cus]
        act = sub_info['资金账号/对手方账号'].tolist()
        act = ["'" + i for i in act]
        sub_info['资金账号/对手方账号'] = act
        #sht.range("A3").value = sub_info.set_index('产品代码')

        #管理人
        mng_lst = sub_info['管理人'].tolist()
        mng_ref = []
        print('管理人', mng_lst)
        for mn in mng_lst:
            #print('管理人',mn)
            if mn == "浙江九章-宁波幻方量化":
                mng = '幻方'
            elif '九章' in mn:
                mng = '九章'
            else:
                mng = '幻方'
            mng_ref.append(mng)
        sub_info['管理人_ref'] = mng_ref
        for mn in list(set(mng_ref)):
            surfix = cus[:4] + '' + mn
            sub_sub_info = sub_info[sub_info['管理人_ref'] == mn]
            sub_sub_info = sub_sub_info.drop(['管理人_ref'], axis=1)
            code_lst = sub_sub_info['划款摘要']
            codes = [code[-6:] for code in code_lst.tolist()]
            if cus == "中信建投证券":
                # code_lst = sub_sub_info['划款摘要']
                # codes = [code[-6:] for code in code_lst.tolist()]
                ops_tps = []
                for code in codes:
                    if int(code) >= 600000:
                        ops_tps.append("网下新股业务-上海")
                    else:
                        ops_tps.append("网下新股业务-深圳")
                sub_sub_info['划款类型'] = ops_tps
                # 打开模板获取表头
                wb = xw.Book(os.path.join(os.path.dirname(this_file), '模板', '中信建投-模板.xlsx'))
                sht = wb.sheets['Sheet1']
                sht.range("A3").value = sub_sub_info.set_index('产品代码')[['经纪商/账户名称','资金账号/对手方账号',
                                                                        '开户行名称','划款类型','划款金额','划款摘要']]
                wb.save(os.path.join(output, filename.strip('.xlsx') + ' '+ surfix + '.xlsx'))
                wb.close()
            elif cus == "国信证券":
                date_lst = sub_sub_info['划款日期'].tolist()
                d_lst = [dt.date().strftime("%Y%m%d") for dt in date_lst]
                sub_sub_info['划款日期'] = d_lst
                sub_sub_info['中签股票'] = codes
                qty = sub_sub_info['获配数量'].tolist()
                qty_to_str = ["'" + str(qt) for qt in qty]
                sub_sub_info['获配数量'] = qty_to_str
                new_info = sub_sub_info[['划款日期','产品名称','中签股票','获配数量','划款金额']]
                new_info.columns = new_info.iloc[0]
                new_info = new_info.iloc[1:].set_index(list(new_info)[0])
                # new_info.columns = range(new_info.shape[1])
                wb = xw.Book(os.path.join(os.path.dirname(this_file), '模板','国信-模板.xls'))
                sht = wb.sheets['Sheet1']
                sht.range("A3").value = new_info
                wb.save(os.path.join(output, filename.strip('.xlsx') + ' '+ surfix + '.xls'))
                wb.close()
            elif cus == "国泰君安证券":
                sub_sub_info['中签股票代码'] = codes
                sub_sub_info['股东代码'] = [code[:-6] for code in code_lst]
                new_info = sub_sub_info[['产品代码','中签股票代码','股东代码','划款金额']]
                new_info.columns = new_info.iloc[0]
                new_info = new_info.iloc[1:].set_index(list(new_info)[0])
                wb = xw.Book(os.path.join(os.path.dirname(this_file), '模板', '国泰君安-模板.xlsx'))
                sht = wb.sheets['Sheet1']
                sht.range("A4").value = new_info
                wb.save(os.path.join(output, filename.strip('.xlsx') + ' '+ surfix + '.xlsx'))
                wb.close()
            else:
                # 打开模板获取表头
                wb = xw.Book(os.path.join(os.path.dirname(this_file), '模板','模板.xlsx'))
                sht = wb.sheets['Sheet1']
                sht.range("A3").value = sub_sub_info.set_index('产品代码')
                wb.save(os.path.join(output, filename.strip('.xlsx') + ' '+ surfix + '.xlsx'))
                wb.close()
