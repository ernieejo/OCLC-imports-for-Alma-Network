import pandas as pd
from pandas import DataFrame, merge, ExcelWriter

#Read the Alma Analytics output
AlmaExport = pd.read_excel('C://Users//...//OCLC Import Script//OCLC-doublecheck.xlsx')
AlmaExport.columns =["MMS IZ","OCLC Control Number (035a)","OCLC Control Number (035z)","001"]

#make a dataframe from the Alma Analytic
df1 = pd.DataFrame(AlmaExport, columns= ['MMS IZ','OCLC Control Number (035a)','OCLC Control Number (035z)','001'])
df1['001']= df1['001'].astype(str)
df1['MMS IZ']= df1['MMS IZ'].astype(str)
print(df1)

#Read the DIFF tab of the comparison_IZ file created with script1
IZ_import = pd.read_excel('C://Users//...//OCLC Import Script//comparison_fileIZ.xlsx', sheet_name='DIFF')
IZ_import.columns = ["a", "b", "MMS IZ", "Existing OCLC number", "035 $a","action"]

#make a dataframe from the comparison_IZ file and set the format as text for the MMS ID. Drop any rows where the action is unresolved or no action
df2 = pd.DataFrame(IZ_import, columns= ['MMS IZ','Existing OCLC number','035 $a','action'])
df2['MMS IZ']= df2['MMS IZ'].astype(str)
df2.drop(df2.index[df2['action'] == 'unresolved'], inplace=True)
df2.drop(df2.index[df2['action'] == 'no action'], inplace=True)
print(df2)

#merge the Alma Analytic and the comparison_fileIZ
df3 = df1.merge(df2, on='MMS IZ', how='inner')
df3['MMS IZ']= df3['MMS IZ'].astype(str)
df3['001']= df1['001'].astype(str)
print(df3)

#Define the sheets that we will have in the comparison_fileNZ
DIFF = df3[df3['035 $a'] != df3['OCLC Control Number (035a)']]
SAME = df3[df3['035 $a'] == df3['OCLC Control Number (035a)']]
a_to_z = DIFF[DIFF['OCLC Control Number (035a)'] != DIFF['OCLC Control Number (035z)']]
no_a_to_z = DIFF[DIFF['035 $a'] == DIFF['OCLC Control Number (035z)']]

print(DIFF)
print(SAME)
print(a_to_z)
print(no_a_to_z)

#Create the comparison_fileNZ file
writer = pd.ExcelWriter('C://Users//...//OCLC Import Script//comparison_fileNZ.xlsx', engine='xlsxwriter')
DIFF.to_excel(writer, sheet_name='DIFF')
SAME.to_excel(writer, sheet_name='SAME')
a_to_z.to_excel(writer, sheet_name='a_to_z')
no_a_to_z.to_excel(writer, sheet_name='no_a_to_z')
writer.save()







