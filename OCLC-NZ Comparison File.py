import pandas as pd
from pandas import DataFrame, merge, ExcelWriter
#Read the Alma Analytics output
AlmaExport = pd.read_excel('OCLC-doublecheck.xlsx', dtype=str)
AlmaExport.columns =["MMS_IZ","OCLC Control Number (035a)","OCLC Control Number (035z)","001"]

#make a dataframe from the Alma Analytic
df1 = pd.DataFrame(AlmaExport, columns= ['MMS_IZ','OCLC Control Number (035a)','OCLC Control Number (035z)','001'])
print(df1)

#Read the DIFF tab of the comparison_IZ file created with script1
IZ_import = pd.read_excel('comparison_file_IZ.xlsx', sheet_name='DIFF', dtype=str)
IZ_import.columns = ["a", "MMS_IZ", "Reported OCLC number", "035 $a","action"]

#Dataframe for do not change list 
values =  pd.read_excel('Do_not_change.xlsx', sheet_name='Do_not_change', dtype=str)
values.columns = ["MMS_IZ"]
values_df = pd.DataFrame(values, columns = ['MMS_IZ'])
values_df['MMS_IZ']= values_df['MMS_IZ'].astype(str)
print(values_df)

#make a dataframe from the comparison_IZ file and set the format as text for the MMS ID. Drop any rows where the action is unresolved or no action
df2 = pd.DataFrame(IZ_import, columns= ['MMS_IZ','Reported OCLC number','035 $a','action'])
df2['MMS_IZ']= df2['MMS_IZ'].astype(str)
print(df2)

#merge the Alma Analytic and the comparison_fileIZ
df3 = df1.merge(df2, on='MMS_IZ', how='outer')
print(df3)


#Define the sheets that we will have in the comparison_fileNZ
do_not_change = df3.merge(values_df, on='MMS_IZ', how='inner')
print(do_not_change)

diff = df3[df3['035 $a'] != df3['OCLC Control Number (035a)']]
updated = df3[df3['035 $a'] == df3['OCLC Control Number (035a)']]
print(diff)
print(updated)

#Create the comparison_fileNZ file
writer = pd.ExcelWriter('comparison_fileNZ.xlsx', engine='xlsxwriter')
do_not_change.to_excel(writer, sheet_name='do_not_change')
diff.to_excel(writer, sheet_name='diff')
updated.to_excel(writer, sheet_name='updated')
writer.save()
