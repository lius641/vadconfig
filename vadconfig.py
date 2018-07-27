import pandas as pd
import re
import numpy as np
with pd.ExcelFile('C:\\Users\\liusi\\Documents\\Data\\YO28762_Query_Status.xls') as xls:
    df1 = pd.read_excel(xls, 0)
    df1 = df1[1:100]
    df3 = df1[df1['Unnamed: 18'].astype(str).str.contains('day|\d', flags=re.IGNORECASE, regex=True)]
    df3.to_csv(path_or_buf='c:\\users\\liusi\\documents\\data\\yo28762_query_status.csv', header = False, index = False)
clean_csv = pd.read_csv(filepath_or_buffer='c:\\users\\liusi\\documents\\data\\yo28762_query_status.csv', nrows = 20)
clean_csv['Date Query Opened'] = pd.to_datetime(clean_csv['Date Query Opened'].str.replace('\\r\\n','')).dt.date
clean_csv['Date Query Answered?'] = pd.to_datetime(clean_csv['Date Query Answered?'].str.replace('\\r\\n','')).dt.date
clean_csv['Date Query Closed'] = pd.to_datetime(clean_csv['Date Query Closed'].str.replace('\\r\\n','')).dt.date
clean_csv['Query Turn Around Days'] = (clean_csv['Date Query Answered?'] - clean_csv['Date Query Opened']).dt.days
#s = clean_csv.groupby('CRF Page').agg({'CRF Page':np.size}).reset_index()
clean_csv['CRF Page Count'] = clean_csv.groupby('CRF Page')['CRF Page'].transform(np.size)


import pandas as pd
import os.path
import numpy as np
path = 'C:\\Users\\liusi\\Documents\\ClinicalProgrammer\\Transformation\\GDSR'
pathin = os.path.join(path,'gds-latest-sdtmv.xlsx')
pathout = os.path.join('C:\\Users\\liusi\\Documents\\Data', 'sdtm_spec.csv')
sasout = os.path.join('C:\\Users\\liusi\\Documents\\Data', 'temp.sas')
df = pd.read_excel(
    pathin, 
    sheet_name = 'DM', 
    header = 6, 
    usecols = [2,3,4,5,6,7], 
    index_col = None).dropna(subset = ['Length'])
df.to_csv(pathout, index = False)
df['Types'] = df['Type'].astype(str)
df['Types'] = np.where(df['Types']=='Num','','$')
df['Lengths'] = df['Length'].astype(str)
df['Lengths'] = df['Lengths'].str.replace('\.0','.')
f = open(sasout,"w")
f.writelines('%macro attributes_ae();\n')
f.writelines('attrib '+ df['Name'] + ' label = "' + df['Label'] + '" length ' + df['Types'] + df['Lengths'] + ';\n')
f.writelines('%mend attributes_ae();')
f.close()