import pandas as pd
from pandas import DataFrame, ExcelWriter
import csv

#Read the BIB processing report
data = pd.read_csv('BibProcessingReport.txt', sep="|", header=None, dtype=str)
data.columns = ["JobID", "MMSID", "Existing 035a", "Incoming 035a", "Action"]

#make a dataframe from the BIB processing report and set the format as text for columns b and c
df = pd.DataFrame(data, columns= ['JobID', 'MMSID', 'Existing 035a', 'Incoming 035a', 'Action'])
df['MMSID']= df['MMSID'].astype(str)

#Define the sheets that we will have in the comparison_fileIZ
DIFF = df[df['Existing 035a'] != df['Incoming 035a']]
SAME = df[df['Existing 035a'] == df['Incoming 035a']]

print(DIFF)
print(SAME)

#Create the comparison_fileIZ file
writer = pd.ExcelWriter('comparison_file.xlsx', engine='xlsxwriter')
DIFF.to_excel(writer, index=False, sheet_name='DIFF')
SAME.to_excel(writer, index=False, sheet_name='SAME')
workbook = writer.book
worksheet = writer.sheets['DIFF']
text_fmt = workbook.add_format({'num_format': '@'})
worksheet.set_column('B:B',20, text_fmt)
writer.save()








