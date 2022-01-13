"""Connect to published Power BI datasets and export data with Python
Author:
    https://LittleBigFrog.xyz - 13.01.2022
"""

import adodbapi #Install via pip install pywin32 or conda install pywin32
import pandas as pd

def GetData(DatasetId,TableQuery):
    ConnectionString = f"""Provider=MSOLAP.8;
    Integrated Security=ClaimsToken;
    Persist Security Info=True;
    Initial Catalog=sobe_wowvirtualserver-{DatasetId};
    Data Source=pbiazure://api.powerbi.com;
    MDX Compatibility=1;
    Safety Options=2;
    MDX Missing Member Mode=Error;
    Identity Provider=https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, 929d0ec0-7a41-4b1e-bc7c-b754a28bddcc;
    Update Isolation Level=2"""
    Query = TableQuery if TableQuery.lower().startswith("select") else  f"evaluate {TableQuery}" 
    conn = adodbapi.connect(ConnectionString)
    DataFrame=pd.read_sql(Query, conn)
    return DataFrame

# Find your DatasetId
# go to https://app.powerbi.com/datahub/datasets
# click on your dataset: end of url is your DatasetId
# click on "Show tables" to see available tables
DatasetId = "ee30a2fc-aezb-4eb8-81d2-d1ec21eec060"
QueryTablesList="Select TABLE_NAME from $SYSTEM.DBSCHEMA_TABLES where TABLE_SCHEMA='MODEL' and TABLE_TYPE='SYSTEM TABLE' "
TableName = "DatasetTableName" #not case sensitive


# List of Dataset tables
AvailableTables=GetData(DatasetId,QueryTablesList)['TABLE_NAME'].tolist()
# Get pandas dataframe for specific table
Data=GetData(DatasetId,TableName)
# Get data from a DAX query
Dax='SUMMARIZE (Sales,Sales[Color],"Sales", SUM ( Sales[Amount] ))'
DataDax=GetData(DatasetId,Dax)
# Export to Excel (requires openpyxl)
Data.to_excel("PowerBIexport.xlsx", index = False)
# Export to csv 
Data.to_csv('out.csv')


# Export table to SQL server
from sqlalchemy import create_engine
#Build your connection: https://docs.sqlalchemy.org/en/14/dialects/mssql.html#module-sqlalchemy.dialects.mssql.pyodbc
connection_uri = create_engine(
    "mssql+pyodbc://scott:tiger@myhost:49242/databasename"
    "?driver=ODBC+Driver+17+for+SQL+Server"
    "&authentication=ActiveDirectoryIntegrated"
)
engine = create_engine(connection_uri, fast_executemany=True)
Data.to_sql("SQL_Table_Name", engine, if_exists="replace", index=False)