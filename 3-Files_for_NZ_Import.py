import pandas as pd

# Read the Excel file
df = pd.read_excel('comparison_fileNZ.xlsx', sheet_name='diff', dtype=str)

# Create a text file for column '001' with header 'MMS ID'
df['001'].to_csv('a_to_z_MMSid_for_set.txt', header=['MMS ID'], index=False)

# Create a new DataFrame for 'for_import_to_NZ' with columns '001' and '035 $a'
for_import_to_NZ = df[['001', '035 $a']]

# Save the new DataFrame to an Excel file with text format
with pd.ExcelWriter('for_import_to_NZ.xlsx', engine='xlsxwriter') as writer:
    for_import_to_NZ.to_excel(writer, index=False, sheet_name='for_import_to_NZ')
    
    # Get the xlsxwriter workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['for_import_to_NZ']
    
    # Set the text format for the entire sheet
    text_fmt = workbook.add_format({'num_format': '@'})
    worksheet.set_column('A:XFD', None, text_fmt)  # Set the entire sheet to text format

    
