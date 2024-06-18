import pandas as pd

etr = pd.read_excel(r"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\ETR\ETR_output.xlsx")
query = pd.read_excel(r"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\SP QUERY\sp_query_cleaned.xlsx")

output = etr.merge(query, how = 'left', left_on='Broker Ref. No.', right_on='Entry No.')
output['DXC'] = output['DXC'].dt.strftime('%m/%d/%Y')
output = output.rename(columns={'Notes_x':'Notes_do','Notes_y':'Notes_sp'})
#output.to_excel('combined_output.xlsx', index=False)

writer = pd.ExcelWriter(r"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\SP QUERY\combined_output.xlsx", engine='xlsxwriter') 
output.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)

workbook = writer.book
worksheet = writer.sheets['Sheet1']

(max_row, max_col) = output.shape

column_settings = [{'header': column} for column in output.columns]

worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

worksheet.set_column(0, max_col - 1, 12)

writer.save()