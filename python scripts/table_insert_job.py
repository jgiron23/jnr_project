import pandas as pd
import pandasql as psql
import pyodbc
import glob, os
from plyer import notification

cnx = pyodbc.connect('Driver={SQL Server};Server=.;Database=JNR;Trusted_Conneciton=Yes')

path = 'C:\\Users\\Juan Giron\\Documents\\git_repo\\Inventory\\'


all_files = glob.glob(os.path.join(path , "*.xlsx"))

li = []

for filename in all_files:
    df = pd.read_excel(filename)
    df['OBJECT_NAME'] = filename.replace(path+"\\","")
    df['OBJECT_NAME'] = filename.replace(path,"")
    df['TOTAL'] = df['TOTAL'].astype(int)
    li.append(df)

frame = pd.concat(li, axis=0, ignore_index=True)
frame = psql.sqldf('''
select
DESCRIPTION,
SKU,
cast(TOTAL as int) as TOTAL,
OBJECT_NAME
from frame
''')

cursor = cnx.cursor()
# Insert Dataframe into SQL Server:
for index, row in frame.iterrows():
     cursor.execute("INSERT INTO dbo.tblRJTB_INVENTORY_SOURCE values(?,?,?,?)", row.DESCRIPTION,row.SKU,row.TOTAL,row.OBJECT_NAME)
cnx.commit()
cursor.close()

notification.notify(
    title = 'SQL Server has been updated!',
    message = 'Updates have been successful',
    timeout = 30
)