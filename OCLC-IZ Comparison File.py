import pandas as pd
from pandas import DataFrame, ExcelWriter
import csv
import xlsxwriter

#Read the BIB processing report
data = pd.read_csv('C://.../BIB_processing_report.txt', sep="|", header=None, dtype=str)
data.columns = ["a", "b", "c", "d", "e"]

#make a dataframe from the BIB processing report and set the format as text for columns b and c
df = pd.DataFrame(data, columns= ['a', 'b', 'c', 'd', 'e'])

#Define the sheets that we will have in the comparison_fileIZ
DIFF = df[df['c'] != df['d']]
SAME = df[df['c'] == df['d']]

print(DIFF)
print(SAME)

#Create the comparison_fileIZ file
writer = pd.ExcelWriter('C://...OCLC Import Script//comparison_fileIZ.xlsx', engine='xlsxwriter')
DIFF.to_excel(writer, index=False, sheet_name='DIFF')
SAME.to_excel(writer, index=False, sheet_name='SAME')
workbook = writer.book
worksheet = writer.sheets['DIFF']
text_fmt = workbook.add_format({'num_format': '@'})
worksheet.set_column('B:B',20, text_fmt)
writer.save()







