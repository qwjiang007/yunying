import pandas as pd
import pymysql
import os
import xlwings as xw

def zs_trans(zs_df):
    mgn_ref = ['幻方' if cus == '宁波幻方量化投资管理合伙企业（有限合伙）' or cus == '浙江九章-宁波幻方量化'
               else '九章' for cus in zs_df['管理人'].tolist()]
    zs_df['管理人_ref'] = mgn_ref
    zs_df['产品代码'] = zs_df['备案编码'].tolist()
    zs_df['划款日期'] = [dt.replace('-','/')] * len(zs_df)
    #ops_tp_lst = list(set(zs_df['划款类型']))
    mgn_lst = list(set(zs_df['管理人_ref']))
    for mgn in mgn_lst:
        #print(mgn)
        raw_df = zs_df[zs_df['管理人_ref'] == mgn]
        ops_tp_lst = list(set(raw_df['划款类型']))
        for tp in ops_tp_lst:
            if tp in ['期转银','银转证']:
                file_path = os.path.join(zs_path, '大跌')
            else:
                file_path = os.path.join(zs_path, '大涨')

            sub_raw_df = raw_df[raw_df['划款类型'] == tp]
            #sub_raw_df['划款日期'] = [dt.replace('-','/') * len(sub_raw_df)]
            #print(sub_raw_df['划款日期'])
            wb = xw.Book(os.path.join(zs_path, '模板.xlsx'))
            sht = wb.sheets['Sheet1']
            sht.range('A3').value = sub_raw_df[['产品代码','经纪商/账户名称','资金账号/对手方账号', '开户行名称','划款类型',
                                                '划款金额', '划款日期', '划款摘要']].set_index('产品代码')
            wb.save(os.path.join(file_path, tp + '-招商' + mgn + ' ' + dt + '.xlsx'))
            wb.close()
            #print('done')

raw_info = pd.read_excel('E:\PythonCode\运营出入金\保证金调整/警告信息.xlsx')['警告信息'].tolist()
errors = []
code_list = []
ft_broker_list = []
reg_num_list = []
cus_list = []
sk_broker_list = []
ft_num_list = []
sk_num_list = []
bank_title_list = []
ops_tp_ft_list = []
ops_tp_sk_list = []
amt_list = []
mng_list = []
for info in raw_info:
    info_lst = info.split()
    code_ref = info_lst[1]
    amt = int(info_lst[3]) * 10000
    #code_lst = code_ref.split("_")
    #code = code_lst[-1].strip("TS").strip('CM').strip("CIT").strip('信用')
    con = pymysql.connect(host='fund.high-flyer.cn',
                          db='基金数据库',
                          user='qw.jiang',
                          password='jqiwei2173192',
                          charset='utf8')
    with con.cursor(cursor=pymysql.cursors.DictCursor) as cur:
        sql = 'SELECT * FROM 资金账户匹配 WHERE 终端名称 = "{}"'.format(code_ref)
        cur.execute(sql)
        res = cur.fetchall()
    if not res:
        errors.append({'终端名': code_ref, '错误': '终端名未查询到相关信息'})
        continue
    else:
        act_info = res[0]
    code_list.append(act_info['终端名后五位'])
    ft_broker_list.append(act_info['期货户经纪商'])
    sk_broker_list.append(act_info['证券户经纪商'])
    cus_list.append(act_info['托管人'])
    reg_num_list.append(act_info['备案编码'])
    ft_num_list.append(act_info['期货户资金账号'])
    sk_num_list.append(act_info['证券户资金账号'])
    bank_title_list.append(act_info['托管账户名称'])
    if amt > 0:
        ops_tp_ft = '期转银'
        ops_tp_sk = '银转证'
    else:
        ops_tp_ft = '银转期'
        ops_tp_sk = '证转银'
    ops_tp_ft_list.append(ops_tp_ft)
    ops_tp_sk_list.append(ops_tp_sk)
    amt_list.append(abs(amt))
    mng_list.append(act_info['管理人'])
ft_df = pd.DataFrame({
    '产品编码': code_list,
    '托管平台': cus_list,
    '备案编码': reg_num_list,
    '经纪商/账户名称': ft_broker_list,
    '资金账号/对手方账号': ft_num_list,
    '开户行名称': bank_title_list,
    '划款类型': ops_tp_ft_list,
    '划款金额': amt_list,
    '划款摘要': ft_num_list,
    '管理人': mng_list
})

sk_df = pd.DataFrame({
    '产品编码': code_list,
    '托管平台': cus_list,
    '备案编码': reg_num_list,
    '经纪商/账户名称': sk_broker_list,
    '资金账号/对手方账号': sk_num_list,
    '开户行名称': bank_title_list,
    '划款类型': ops_tp_sk_list,
    '划款金额': amt_list,
    '划款摘要': sk_num_list,
    '管理人': mng_list
})

dt = str(pd.datetime.today().date())
writer = pd.ExcelWriter('E:\PythonCode\运营出入金\保证金调整/保证金调整 {}.xlsx'.format(dt),engine='xlsxwriter')
ft_df.set_index('产品编码').to_excel(writer, sheet_name='银期')
sk_df.set_index('产品编码').to_excel(writer, sheet_name='银证')
writer.save()

zs_path = "E:\PythonCode\运营出入金\保证金调整\期货保证金调整-招商"
zs_ft_df = ft_df[ft_df['托管平台'] == '招商证券']
zs_sk_df = sk_df[sk_df['托管平台'] == '招商证券']

zs_trans(zs_ft_df)
zs_trans(zs_sk_df)

if errors:
    print('#############################存在报错！##############################')
    for row in errors:
        print(row)
    pd.DataFrame(errors).set_index('终端名').to_excel(os.path.join(zs_path,'..','报错信息.xlsx'))
else:
    print('''
    ----------------------------------
            文件已全部生成
    ----------------------------------''')



