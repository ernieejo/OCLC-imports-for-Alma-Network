import pandas as pd
from pandas import DataFrame, merge, ExcelWriter
#Read the Alma Analytics output
AlmaExport = pd.read_excel('C://...OCLC Import Script//OCLC-doublecheck.xlsx', dtype=str)
AlmaExport.columns =["MMS_IZ","OCLC Control Number (035a)","OCLC Control Number (035z)","001"]

#make a dataframe from the Alma Analytic
df1 = pd.DataFrame(AlmaExport, columns= ['MMS_IZ','OCLC Control Number (035a)','OCLC Control Number (035z)','001'])
print(df1)

#Read the DIFF tab of the comparison_IZ file created with script1
IZ_import = pd.read_excel('C://...OCLC Import Script//comparison_fileIZ.xlsx', sheet_name='DIFF', dtype=str)
IZ_import.columns = ["a", "MMS_IZ", "Existing OCLC number", "035 $a","action"]

#Dataframe for do not change list 
values =  pd.read_excel('C://...OCLC Import Script//Do_not_change.xlsx', sheet_name='Do_not_change', dtype=str)
values.columns = ["MMS_IZ"]
values_df = pd.DataFrame(values, columns = ['MMS_IZ'])
values_df['MMS_IZ']= values_df['MMS_IZ'].astype(str)
print(values_df)

#make a dataframe from the comparison_IZ file and set the format as text for the MMS ID. Drop any rows where the action is unresolved or no action
df2 = pd.DataFrame(IZ_import, columns= ['MMS_IZ','Existing OCLC number','035 $a','action'])
df2['MMS_IZ']= df2['MMS_IZ'].astype(str)
df2.drop(df2.index[df2['action'] == 'unresolved'], inplace=True)
df2.drop(df2.index[df2['action'] == 'no action'], inplace=True)
print(df2)

#merge the Alma Analytic and the comparison_fileIZ
df3 = df1.merge(df2, on='MMS_IZ', how='outer')
print(df3)


#Define the sheets that we will have in the comparison_fileNZ
do_not_change = df3.merge(values_df, on='MMS_IZ', how='inner')
print(do_not_change)

diff = df3[df3['035 $a'] != df3['OCLC Control Number (035a)']]
Review = diff[diff['OCLC Control Number (035a)'] == diff['OCLC Control Number (035z)']]
print(diff)
print(Review)

#Create the comparison_fileNZ file
writer = pd.ExcelWriter('C://...OCLC Import Script//comparison_fileNZ.xlsx', engine='xlsxwriter')
do_not_change.to_excel(writer, sheet_name='do_not_change')
diff.to_excel(writer, sheet_name='diff')
Review.to_excel(writer, sheet_name='duplication in 035a-035z')
writer.save()

