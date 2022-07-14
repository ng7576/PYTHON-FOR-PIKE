import pandas as pd;
import numpy as np;
import os;


dirname = "EXCEL";
nameandcountmap = {};
emptySheets ={};


for excelfiles in os.listdir(dirname):
  nameandcountmap[excelfiles[:7]] = 0;

from pandas.io import excel
for excelfiles in os.listdir(dirname):
  if (excelfiles[0]==".") :
    continue

  df_current = pd.read_excel(os.path.join(dirname, excelfiles));
  # print(df_current['Engineer Approval']);
  if df_current['Engineer Approval'].isnull().values.all():
    emptySheets[excelfiles[:7]] = 0;
  else:
    count=0
    for idx, rows in df_current.iterrows():
      # print(rows[6], type(rows[6]));
      if rows[6] == 'Remove Invalid Address' or type(rows[6])==float:
        continue;
      else:
        count+=1;
      nameandcountmap[excelfiles[:7]]=count;       



df_master1 = pd.read_excel('masterSheet.xlsx');
df_master = df_master1;
df_master['F2 CFAS'].str.strip();
df_master['NUMBER OF ENTRIES']=0;
df_master['Empty Sheet'] = ' ';



count=0;
for idx, rows in df_master.iterrows():
  if (rows[5] in emptySheets):
    df_master.at[idx, 'Empty Sheet'] = 'This CFAS is Empty in BIZ portal';
 
  try :
    if rows[5] in nameandcountmap:    
      val =nameandcountmap[rows[5]];
      df_master.at[idx, 'NUMBER OF ENTRIES'] = val;
  except:
    print(rows[5],"is not in hashmap", count);
    count +=1;


df_master.to_excel("Final.xlsx");
    