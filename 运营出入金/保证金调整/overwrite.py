from tools.db_connect import get_fund_db
import pandas as pd

raw_info = pd.read_excel('E:\PythonCode\运营出入金\保证金调整/资金账户匹配.xls', dtype=str)
with get_fund_db() as cur:
    sql = "DELETE FROM 资金账户匹配"
    cur.execute(sql)
    print("已删除{}条记录".format(cur.rowcount))
count = 0
for ind, row in raw_info.iterrows():
    key_lst = tuple(row.index)
    value_lst = tuple(v if v!='nan' else "无" for v in row)
    #values = tuple(["%s"]* len(value_lst))
    with get_fund_db() as cur:
        sql = '''
        INSERT INTO 资金账户匹配 (终端名后五位, 产品名称, 备案编码, 管理人, 托管人, 托管账户名称, 托管账户开户行, 托管银行账号, 证券户经纪商, 证券户资金账号, 客户代码, 期货户经纪商, 期货户资金账号, 终端名称) VALUES{}
        '''.format(value_lst)
        cur.execute(sql)
        count += cur.rowcount
print('更新完毕，加入{}条记录'.format(count))
    #print(cur._executed)
    #print(cur.rowcount)