from tools.db_connect import get_fund_db
from tools.update_db import update_DB
import pandas as pd
import time
import pymysql

con = pymysql.connect(host='fund.high-flyer.cn',
                      db='基金数据库',
                      user='qw.jiang',
                      password='jqiwei2173192',
                      charset='utf8')
cur = con.cursor(cursor=pymysql.cursors.DictCursor)

raw_info = pd.read_excel('E:\PythonCode\运营出入金\保证金调整/账户信息.xlsx', dtype=str)

# code_lst = [str(100000 + code)[1:] for code in raw_info['产品编号']]
# raw_info['终端名后五位'] = code_lst
# prod_lst = raw_info['产品名称'].tolist()
# reg_num_lst = raw_info['备案编码'].tolist()
# mng_lst = raw_info['管理人'].tolist()
# cus_lst = raw_info['托管人'].tolist()
cols = ','.join(raw_info.columns.tolist())

for i, row in raw_info.iterrows():
    update_lst = []
    for key, value in zip(raw_info.columns.tolist(), list(row)):
        if key == "终端名后五位":
            continue
        else:
            update_lst.append(key + '=' + '"' + value + '"')
    update = ', '.join(update_lst)
    sql = '''INSERT INTO 资金账户匹配 ({}) VALUES {}
    ON DUPLICATE KEY UPDATE {}
    '''.format(cols, tuple(row), update)
    cur.execute(sql)
    print(cur._executed)
    print(cur.rowcount)
    con.commit()
    #con.close()
    time.sleep(0.2)