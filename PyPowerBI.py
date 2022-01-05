# adodpapi to be installed via  pip install pywin32 or conda install pywin32
import adodbapi 
import pandas as pd

DatasetId = ""
TableName = ""
MyConnectionString = f"Provider=MSOLAP.8;Integrated Security=ClaimsToken;Persist Security Info=True;Initial Catalog=sobe_wowvirtualserver-{DatasetId};Data Source=pbiazure://api.powerbi.com;MDX Compatibility=1;Safety Options=2;MDX Missing Member Mode=Error;Identity Provider=https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, 929d0ec0-7a41-4b1e-bc7c-b754a28bddcc;Update Isolation Level=2"
MyQuery = f"evaluate {TableName}"
conn = adodbapi.connect(MyConnectionString)
df = pd.read_sql(MyQuery, conn)
print(df)