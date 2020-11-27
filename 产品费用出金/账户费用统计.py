#用于统计月初各产品需要往托管户入金金额
from tools.db_connect import get_fund_db
import pandas as pd
import os
import yaml
import numpy as np


this_file = os.path.realpath(__file__)
#this_file = "E:/PythonCode/运营出入金/产品费用出金/"
cus_balance_conf = os.path.join(os.path.dirname(this_file), '..', '..','configs', 'cus_balance_config.yml')
with open(cus_balance_conf, 'r', encoding='utf8') as conf:
    cus_balance_configs = yaml.safe_load(conf)

def parse_income(income_df):
    head = list(income_df.columns)
    fee_lst = ['管理费','业绩报酬','赎回费']
    ref = list(filter(lambda x: x in head, fee_lst ))[0]
    #print(ref)
    cus_ref = list(filter(lambda x: x in head, ['宁波','九章']))[0]
    #end_ind = list(income_df[pd.isnull(income_df[cus_ref])].index)[-2]
    sep_ind = list(income_df[~income_df.applymap(np.isreal)[ref]].index)
    income_df = income_df.rename(columns={cus_ref: '产品名称'})
    income1 = income_df[: sep_ind[0] - 1]
    income1 = income1.rename(columns={'实际到账': '实际到账' + ref })
    income2 = income_df[sep_ind[0]: sep_ind[1] - 1]
    income2_fee = income2[ref].tolist()[0]
    income2 = income2[1:]
    income2 = income2.rename(columns={ref: income2_fee, '实际到账': '实际到账' + income2_fee})
    #income3 = income_df[sep_ind[1]:sep_ind[2]-1]
    income3 = income_df[sep_ind[1]: len(income_df) -1]
    income3_fee = income3[ref].tolist()[0]
    income3 = income3[1:]
    income3 = income3.rename(columns={ref: income3_fee, '实际到账': '实际到账' + income3_fee})
    income_res = income1.merge(income2, how='outer',
                               on=['产品名称']).merge(income3, how='outer', on=['产品名称'])
    income_res = income_res.drop_duplicates(subset='产品名称', keep='first').fillna(0)
    return income_res


#获取增值税费
def parse_tax(income_res, tax_df):
    #sep_ind = list(tax_df[tax_df.applymap(np.isreal)['产品名称']].index)
    end_ind = list(tax_df[tax_df['产品名称'] == '合计：'].index)[0]
    tax_df = tax_df[: end_ind]
    tax_df = tax_df.drop_duplicates(subset='产品名称', keep='first')
    return income_res.merge(tax_df, how= 'outer', on=['产品名称']).drop_duplicates(subset='产品名称', keep='first').fillna(0)

#计算托管户需支付费用
def fee_cal(income_res):
    income_res = income_res.replace('暂缺',0)
    #需支付管理费
    mgn_fee = np.array(income_res['管理费']) - np.array(income_res['实际到账管理费'])
    #需支付赎回费
    red_fee = np.array(income_res['赎回费']) - np.array(income_res['实际到账赎回费'])
    #需支付业绩报酬
    perf_fee = np.array(income_res['业绩报酬']) - np.array(income_res['实际到账业绩报酬'])
    #需支付税费
    tax = np.array(income_res['应到账税费']) - np.array(income_res['实际到账税费'])
    #总计托管户需支付
    sub_total = mgn_fee + red_fee + perf_fee + tax
    income_res['需支付管理费'] = mgn_fee
    income_res['需支付赎回费'] = red_fee
    income_res['需支付业绩报酬'] = perf_fee
    income_res['需支付增值税费'] = tax
    income_res['总计托管户需支付'] = sub_total
    return income_res

#爬取托管户余额
def cus_act_balance(income_res):
    balance_folder = os.path.join(os.path.dirname(this_file),'源文件','托管户余额')
    file_lst = os.listdir(balance_folder)
    balance_df = pd.DataFrame()
    for filename in file_lst:
        sub_balance_df = pd.DataFrame()
        if "$" in filename:
            continue
        else:
            if '招商' in filename:
                balance_config = cus_balance_configs['zx']
                sub_balance_df = pd.read_excel(balance_folder + '/' + filename)[[balance_config['prod_name'],balance_config['balance']]].rename(columns={'余额(元)': '托管户余额'})
                sub_balance_df['托管户余额'] = sub_balance_df['托管户余额'].str.replace(',', '').astype(float)
        balance_df = pd.concat([balance_df, sub_balance_df], ignore_index=True)
    #爬取完托管户余额文件夹后会生成一个产品托管户余额的df, 合并到income_res中
    income_res = income_res.merge(balance_df, how='left', on=['产品名称'])
    print(income_res)
    #计算托管户需入金

    amt = np.array(income_res['总计托管户需支付']) - np.array(income_res['托管户余额'])
    amt = [i if i >=0 else np.nan if np.isnan(i) else 0 for i in amt]
    income_res['托管户需入金'] = amt
    return income_res
#获取产品信息
def fecth_prod_info(income_res):
    prod_lst = income_res['产品名称'].tolist()
    cus_lst = []
    prod_code = []
    s_broker_lst = []
    s_act_num_lst = []
    f_broker_lst = []
    f_act_num_lst = []
    dbc = get_fund_db
    missing = '未查询到信息'
    for prod in prod_lst:
        with dbc() as cur:
            sql = '''
            SELECT 产品名称, 托管人, 终端名后五位, 证券户经纪商, 证券户资金账号, 期货户经纪商, 期货户资金账号 FROM 资金账户匹配
            WHERE 产品名称 = %s
            '''
            cur.execute(sql, prod)
            print(cur._executed)
            res_act_info = cur.fetchall()

        if res_act_info:
            res_act = res_act_info[0]
            cus = res_act['托管人']
            cus_lst.append(cus)
            res_act = res_act_info[0]
            prod_code.append(res_act['终端名后五位'])
            s_broker_lst.append(res_act['证券户经纪商'])
            s_act_num_lst.append(res_act['证券户资金账号'])
            f_broker_lst.append(res_act['期货户经纪商'])
            f_act_num_lst.append(res_act['期货户资金账号'])
        else:
            print('%s 无法查询到信息，请确认产品名称准确并且在数据库中录入相关信息。'%prod)
            cus_lst.append(missing)
            prod_code.append(missing)
            s_broker_lst.append(missing)
            s_act_num_lst.append(missing)
            f_broker_lst.append(missing)
            f_act_num_lst.append(missing)

    income_res['终端名后五位'] = prod_code
    income_res['托管人'] = cus_lst
    income_res['证券户经纪商'] = s_broker_lst
    income_res['证券户资金账号'] = s_act_num_lst
    income_res['期货户经纪商'] = f_broker_lst
    income_res['期货户资金账号'] = f_act_num_lst


income_file_path = os.path.join(os.path.dirname(this_file), '源文件','产品收入.xlsx')
tax_file_path = os.path.join(os.path.dirname(this_file), '源文件','资管增值税.xlsx')
income_df_jz = pd.read_excel(income_file_path, sheet_name='九章', usecols=[0,1,2]).fillna('暂缺')
income_df_hf = pd.read_excel(income_file_path, sheet_name='量化', usecols=[0,1,2]).fillna('暂缺')
#解析各项收入
income_res_jz = parse_income(income_df_jz)
income_res_hf = parse_income(income_df_hf)

#产品税费
tax_jz = pd.read_excel(tax_file_path, sheet_name='九章', usecols=[0,1,2], skiprows=3).rename(columns={'实际到账': '实际到账税费'}).fillna(0)
tax_hf = pd.read_excel(tax_file_path, sheet_name='量化', usecols=[0,1,2], skiprows=3).rename(columns={'实际到账': '实际到账税费'}).fillna(0)

#合并产品税费
income_res_jz = parse_tax(income_res_jz, tax_jz)
income_res_hf = parse_tax(income_res_hf, tax_hf)

#计算托管户需支付总费用
income_res_jz = fee_cal(income_res_jz.fillna(0))
income_res_hf = fee_cal(income_res_hf.fillna(0))

#爬取托管户余额
income_res_jz = cus_act_balance(income_res_jz)
income_res_hf = cus_act_balance(income_res_hf)
#爬取产品信息
fecth_prod_info(income_res_jz)
fecth_prod_info(income_res_hf)

income_res_jz = income_res_jz.set_index('产品名称').to_excel(os.path.join(os.path.dirname(this_file), '九章-托管户入金.xlsx'))
income_res_hf= income_res_hf.set_index('产品名称').to_excel(os.path.join(os.path.dirname(this_file), '宁波-托管户入金.xlsx'))






