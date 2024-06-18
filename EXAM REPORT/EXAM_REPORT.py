import pandas as pd 
import numpy as np
import datetime

def year_to_date_report(filepath, output):
    #mot dict 
    mot_dict = {40:'Air',30:'Truck',21:'Rail',11:'Full Container',10:'LCL'}

    data = pd.read_excel(filepath)
    #entry_count: count total unique entries 
    entry_count = len(data.drop_duplicates('Entry Number')['Entry Number'])
    #data filtering
    data = data[((data['Manifest Hold Crt.'].isnull() == False) & (data['Manifest Hold Rmv.'].isnull()==False)) | ((data['Manifest Hold Crt.'].isnull() == True) & (data['Manifest Hold Rmv.'].isnull()==True) & (data['Intensive Exam Opened'].isnull()==True) & (data['CBP Hold and Exam Handling'].isnull()==False))]
    data_duplicates = data.reset_index(drop = True)
    data = data.dropna(subset=['CBP Hold and Exam Handling']).drop_duplicates('Entry Number').reset_index(drop=True)
    #examined_count: count unique entries that are examined
    examined_count = len(data['Entry Number'])
    data['MOT Key'] = data['MOT'].apply(lambda x: mot_dict[x])

    ##helper function for finding masterbill 
    def find_master_bill(entry_num):
        filtered = list(data_duplicates[data_duplicates['Entry Number'] == entry_num]['Master Bill'])
        cleaned = [x for x in filtered if pd.isnull(x) == False]
        return cleaned[0]
    
    #sheet 1: examined entries and Masterbill 
    report = pd.DataFrame()
    report['Entry Number'] = data['Entry Number']
    report['MBL'] = report['Entry Number'].apply(lambda x:find_master_bill(x))

    #sheet 2: HTS non duplicate 
    report2 = data_duplicates['Tariff'].value_counts().rename_axis('HTS').reset_index(name = 'count')
    report2 = report2[report2['count'] > 1] 

    #sheet 4: Production Description non duplicate 
    report4 = data_duplicates['Description.1'].value_counts().rename_axis('Product Description').reset_index(name = 'count')
    report4 = report4[report4['count'] > 1]

    #sheet 6: exporter name 
    report6 = data['Exporter Name'].value_counts().rename_axis('Exporter Name').reset_index(name = 'count')

    #sheet 7: importer name 
    report7 = data['Importer'].value_counts().rename_axis('Importer').reset_index(name = 'count')

    #sheet 8: MOT
    report8 = data['MOT Key'].value_counts().rename_axis('MOT').reset_index(name = 'count')

    #sheet 9: Difference between arrival date and cargo release date 
    report9 = pd.DataFrame()
    report9['Entry Number'] = data['Entry Number']
    report9['Arrival Date'] = data['Arrival Date']
    report9['Cargo Release Date'] = data['Cargo Release Date']
    f = lambda x: len(pd.bdate_range(x['Arrival Date'], x['Cargo Release Date']))
    report9['Date Difference (Only Including Busniess Days)'] = report9.apply(f, axis = 1)
    report9['Arrival Date'] = report9['Arrival Date'].dt.date
    report9['Cargo Release Date'] = report9['Cargo Release Date'].dt.date
    report9['MBL'] = report9['Entry Number'].apply(lambda x:find_master_bill(x))

    #sheet 11: exam percentage 
    report11 = pd.DataFrame()
    report11['Total Entries'] = [entry_count]
    report11['Examined Entries'] = [examined_count]
    report11['Exam Ratio'] = [(examined_count/entry_count)*100]

    #writes all sheets onto excel
    writer = pd.ExcelWriter(output)
    report.to_excel(writer, index = False, startrow=1)
    sheet1 = writer.sheets['Sheet1']
    sheet1.write_string(0,0, 'Entries Examined')
    report2.to_excel(writer, index = False, sheet_name='Sheet1', startcol=3, startrow=1)
    sheet1.write_string(0,3, 'Common HTS among examined entries')
    report4.to_excel(writer, index=False, sheet_name = 'Sheet1',startcol=6, startrow=1)
    sheet1.write_string(0,6,'Common Product Description among examined Entries')
    report6.to_excel(writer, index=False, sheet_name = 'Sheet1',startcol=9, startrow=1)
    sheet1.write_string(0,9,'Exporter Name')
    report7.to_excel(writer, index=False, sheet_name = 'Sheet1',startcol=12, startrow=1)
    sheet1.write_string(0,12,'Importer Name')
    report8.to_excel(writer, index=False, sheet_name = 'Sheet1',startcol=15, startrow=1)
    sheet1.write_string(0,15,'MOT')
    report11.to_excel(writer, index=False, sheet_name = 'Sheet1',startcol=18, startrow=1)
    sheet1.write_string(0,18,'Exam Ratios')
    report9.to_excel(writer, index = False, sheet_name= 'Sheet2', startrow=1)
    sheet2 = writer.sheets['Sheet2']
    sheet2.write_string(0,0, 'Date Difference between Arrival Date to Cargo Release Date')
    writer.save()

year_to_date_report(r"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\EXAM_REPORT\CTS_STATUS.xlsx", r"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\EXAM_REPORT\EXAM_REPORT_output.xlsx")