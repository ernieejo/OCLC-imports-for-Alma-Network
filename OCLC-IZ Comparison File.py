import pandas as pd
from pandas import DataFrame, ExcelWriter
import csv

#Read the BIB processing report
data = pd.read_csv('C://...//OCLC Imports//BibProcessingReport.txt', sep="|", header=None)
data.columns = ["a", "b", "c", "d", "e"]

#make a dataframe from the BIB processing report and set the format as text for columns b and c
df = pd.DataFrame(data, columns= ['a', 'b', 'c', 'd', 'e'])
df['b']= df['b'].astype(str)

#Define the sheets that we will have in the comparison_fileIZ
DIFF = df[df['c'] != df['d']]
SAME = df[df['c'] == df['d']]

print(DIFF)
print(SAME)

#Create the comparison_fileIZ file
writer = pd.ExcelWriter('C://...//OCLC Imports//comparison_fileIZ.xlsx', engine='xlsxwriter')
DIFF.to_excel(writer, sheet_name='DIFF')
SAME.to_excel(writer, sheet_name='SAME')
writer.save()








