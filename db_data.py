from sqlalchemy import create_engine
import pandas as pd
#pymssql 在全局安装成功，但在虚拟环境下未成功。。

engine=create_engine('mssql+pymssql://Central_Mfg_User_NMS:Central_Mfg_User_NMS@123@CVCM-SQL01/CVCM')

print(engine.table_names())

sql="SELECT * FROM CG1_POAccrual_Extract WHERE ORG='FOC'"
df=pd.read_sql(sql,engine)
df[:100].to_excel('PO.xlsx',index=False)
print(df.head(100))