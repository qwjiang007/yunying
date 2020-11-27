import pandas as pd
import time
import pymysql

con = pymysql.connect(host='fund.high-flyer.cn',
                      db='基金数据库',
                      user='qw.jiang',
                      password='jqiwei2173192',
                      charset='utf8')
cur = con.cursor(cursor=pymysql.cursors.DictCursor)
update_info = pd.read_excel('E:\PythonCode\运营出入金\保证金调整\终端名修改信息.xlsx')['修改信息'].tolist()

for info in update_info:
    info_lst = info.split()
    sql = '''
    UPDATE 资金账户匹配
    SET `终端名称` = '{}'
    WHERE `终端名称` = '{}'
    '''.format(info_lst[1], info_lst[0])
    try:
        cur.execute(sql)
        print(cur._executed)
        print(cur.rowcount)
        con.commit()
    except Exception as e:
        print(e)
        print('终端名： {} 未成功修改。'.format(info_lst[0]))
    time.sleep(0.2)