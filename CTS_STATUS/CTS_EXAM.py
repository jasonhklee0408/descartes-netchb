import pandas as pd 
import numpy as np
import datetime
from dateutil.relativedelta import relativedelta

def year_to_date_report(filepath, output):
    #mot dict 
    mot_dict = {40:'Air',30:'Truck',21:'Rail',11:'Full Container',10:'LCL'}

    data = pd.read_excel(filepath)
    data1 = data
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
        if len(cleaned) == 0:
            return None
        else:
            return cleaned[0]
    
    ##helper function for checking if date column cells are empty are not 
    def calculate_date_range(arrival_date, cargo_release_date):
        if pd.isnull(arrival_date):
            return ''
        elif pd.isnull(cargo_release_date):
            return ''
        else:
            return len(pd.date_range(arrival_date,cargo_release_date))

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
    #f = lambda x: len(pd.bdate_range(x['Arrival Date'], x['Cargo Release Date']))
    report9['Date Difference (Only Including Busniess Days)'] = report9.apply(lambda x: calculate_date_range(x['Arrival Date'], x['Cargo Release Date']), axis = 1)
    report9['Arrival Date'] = report9['Arrival Date'].dt.date
    report9['Cargo Release Date'] = report9['Cargo Release Date'].dt.date
    report9['MBL'] = report9['Entry Number'].apply(lambda x:find_master_bill(x))

    #sheet 11: exam percentage 
    report11 = pd.DataFrame()
    report11['Total Entries'] = [entry_count]
    report11['Examined Entries'] = [examined_count]
    report11['Exam Ratio'] = [(examined_count/entry_count)*100]

    #Exam in progress 

    #sheet12: entries exam in progress + Date Difference between arrival date and today's date
    report12 = data1[(data1['Release Status'] != 'RLS') & (data1['Cargo Release Date'].isnull()) & (data1['Entry Number'].isnull() == False) & (data1['7501 Status'] != 'INP') & ((data1['Release Status'] == 'INT') | (data1['Release Status'] == 'REV'))]
    report12 = report12[['Entry Number','Importer','Exporter Name','Arrival Date','Manifest Hold Crt.','Entry Port','Master Bill','Container No']]
    report12['Date Difference'] = report12['Arrival Date'].apply(lambda x: abs((x.date() - datetime.datetime.now().date()).days))
    report12 = report12.drop_duplicates(subset=['Entry Number'], keep = 'first')
    report12['Arrival Date'] = report12['Arrival Date'].apply(lambda x: x.date())
    report12['Manifest Hold Crt.'] = report12['Manifest Hold Crt.'].apply(lambda x: x.date())

    #sheet 13: total shipment exam ratios
    today = pd.Timestamp(datetime.datetime.now().date())
    month_3 = pd.Timestamp(datetime.datetime.now().date() - relativedelta(months=+3))
    month_2 = pd.Timestamp(datetime.datetime.now().date() - relativedelta(months=+2))
    month_1 = pd.Timestamp(datetime.datetime.now().date() - relativedelta(months=+1))
    month_range = [month_3,month_2,month_1]

    total_entry_count = []
    examined_entry_count = []
    exam_ratio_count = []
    date_range = []


    for i in month_range:
        monthdf = data1[(data1['Arrival Date'] > i) & (data1['Arrival Date'] <= today)]
        entry_count = len(monthdf.drop_duplicates('Entry Number')['Entry Number'])
        filtered = monthdf[((monthdf['Manifest Hold Crt.'].isnull() == False) & (monthdf['Manifest Hold Rmv.'].isnull()==False)) | 
                        ((monthdf['Manifest Hold Crt.'].isnull() == True) & (monthdf['Manifest Hold Rmv.'].isnull()==True) & (monthdf['Intensive Exam Opened'].isnull()==True) & (monthdf['CBP Hold and Exam Handling'].isnull()==False)) | 
                        ((monthdf['Release Status'] != 'RLS') & (monthdf['Cargo Release Date'].isnull()) & (monthdf['Entry Number'].isnull() == False) & (monthdf['7501 Status'] != 'INP') & ((monthdf['Release Status'] == 'INT') | (monthdf['Release Status'] == 'REV')))].reset_index(drop=True)
        exam_entry = len(filtered.drop_duplicates('Entry Number',keep='first').reset_index(drop=True))
        exam_ratio = (exam_entry/entry_count)
        total_entry_count.append(entry_count)
        examined_entry_count.append(exam_entry)
        exam_ratio_count.append(exam_ratio)
        date_range.append(str(i.date()) + ' to ' + str(today.date()))

    report13 = pd.DataFrame()
    report13['Date Range'] = date_range
    report13['Total Entries'] = total_entry_count
    report13['Examined Entries'] = examined_entry_count
    report13['Exam Ratio'] = exam_ratio_count
    report13['Exam Ratio'] = report13['Exam Ratio'].map('{:.2%}'.format)

    #sheet 14: Exam counts and ratios
    #Product Description
    #** filter pd by dropping hts starting with 9903
    no_9903 = data1[data1['Tariff'].astype(str).str.startswith('9903') == False]
    description_pd = no_9903['Description.1'].value_counts().rename_axis('Product Description').reset_index(name = 'Total Entries')
    filtered = no_9903[((no_9903['Manifest Hold Crt.'].isnull() == False) & (no_9903['Manifest Hold Rmv.'].isnull()==False)) | ((no_9903['Manifest Hold Crt.'].isnull() == True) & (no_9903['Manifest Hold Rmv.'].isnull()==True) & (no_9903['Intensive Exam Opened'].isnull()==True) & (no_9903['CBP Hold and Exam Handling'].isnull()==False)) | ((no_9903['Release Status'] != 'RLS') & (no_9903['Cargo Release Date'].isnull()) & (no_9903['Entry Number'].isnull() == False) & (no_9903['7501 Status'] != 'INP') & ((no_9903['Release Status'] == 'INT') | (no_9903['Release Status'] == 'REV')))]['Description.1'].value_counts().rename_axis('Product Description').reset_index(name = 'Examined Entries')
    description_pd = description_pd.merge(filtered, on='Product Description',how='left')
    description_pd['Examined Entries'] = description_pd['Examined Entries'].fillna(0)
    description_pd['Examined Entries'] = description_pd['Examined Entries'].astype('int64')
    description_pd['Exam Ratio'] = (description_pd['Examined Entries']/description_pd['Total Entries'])
    description_pd = description_pd[description_pd['Exam Ratio'] >= 0.1]
    description_pd['Exam Ratio'] = description_pd['Exam Ratio'].map('{:.2%}'.format)

    #HTS
    tariff_pd = no_9903['Tariff'].value_counts().rename_axis('HTS').reset_index(name = 'Total Entries')
    filtered =no_9903[((no_9903['Manifest Hold Crt.'].isnull() == False) & (no_9903['Manifest Hold Rmv.'].isnull()==False)) | ((no_9903['Manifest Hold Crt.'].isnull() == True) & (no_9903['Manifest Hold Rmv.'].isnull()==True) & (no_9903['Intensive Exam Opened'].isnull()==True) & (no_9903['CBP Hold and Exam Handling'].isnull()==False)) | ((no_9903['Release Status'] != 'RLS') & (no_9903['Cargo Release Date'].isnull()) & (no_9903['Entry Number'].isnull() == False) & (no_9903['7501 Status'] != 'INP') & ((no_9903['Release Status'] == 'INT') | (no_9903['Release Status'] == 'REV')))]['Tariff'].value_counts().rename_axis('HTS').reset_index(name = 'Examined Entries')
    tariff_pd = tariff_pd.merge(filtered, on='HTS',how='left')
    tariff_pd['Examined Entries'] = tariff_pd['Examined Entries'].fillna(0)
    tariff_pd['Examined Entries'] = tariff_pd['Examined Entries'].astype('int64')
    tariff_pd['Exam Ratio'] = (tariff_pd['Examined Entries']/tariff_pd['Total Entries'])
    tariff_pd = tariff_pd[tariff_pd['Exam Ratio'] > 0.1]
    tariff_pd['Exam Ratio'] = tariff_pd['Exam Ratio'].map('{:.2%}'.format)
    tariff_pd = tariff_pd[tariff_pd['Examined Entries'] > 0]


    #Importer
    importer_pd = data1.drop_duplicates('Entry Number')['Importer'].value_counts().rename_axis('Importer').reset_index(name = 'Total Entries')
    filtered = data1[((data1['Manifest Hold Crt.'].isnull() == False) & (data1['Manifest Hold Rmv.'].isnull()==False)) | ((data1['Manifest Hold Crt.'].isnull() == True) & (data1['Manifest Hold Rmv.'].isnull()==True) & (data1['Intensive Exam Opened'].isnull()==True) & (data1['CBP Hold and Exam Handling'].isnull()==False)) | ((data1['Release Status'] != 'RLS') & (data1['Cargo Release Date'].isnull()) & (data1['Entry Number'].isnull() == False) & (data1['7501 Status'] != 'INP') & ((no_9903['Release Status'] == 'INT') | (no_9903['Release Status'] == 'REV')))].drop_duplicates('Entry Number')['Importer'].value_counts().rename_axis('Importer').reset_index(name = 'Examined Entries')
    importer_pd = importer_pd.merge(filtered, on='Importer',how='left')
    importer_pd['Examined Entries'] = importer_pd['Examined Entries'].fillna(0)
    importer_pd['Examined Entries'] = importer_pd['Examined Entries'].astype('int64')
    importer_pd['Exam Ratio'] = (importer_pd['Examined Entries']/importer_pd['Total Entries'])
    importer_pd['Exam Ratio'] = importer_pd['Exam Ratio'].map('{:.2%}'.format)
    importer_pd = importer_pd[importer_pd['Examined Entries'] > 0]

    #Exporter
    exporter_pd = data1.drop_duplicates('Entry Number')['Exporter Name'].value_counts().rename_axis('Exporter').reset_index(name = 'Total Entries')
    filtered = data1[((data1['Manifest Hold Crt.'].isnull() == False) & (data1['Manifest Hold Rmv.'].isnull()==False)) | ((data1['Manifest Hold Crt.'].isnull() == True) & (data1['Manifest Hold Rmv.'].isnull()==True) & (data1['Intensive Exam Opened'].isnull()==True) & (data1['CBP Hold and Exam Handling'].isnull()==False)) | ((data1['Release Status'] != 'RLS') & (data1['Cargo Release Date'].isnull()) & (data1['Entry Number'].isnull() == False) & (data1['7501 Status'] != 'INP') & ((no_9903['Release Status'] == 'INT') | (no_9903['Release Status'] == 'REV')))].drop_duplicates('Entry Number')['Exporter Name'].value_counts().rename_axis('Exporter').reset_index(name = 'Examined Entries')
    exporter_pd = exporter_pd.merge(filtered, on='Exporter',how='left')
    exporter_pd['Examined Entries'] = exporter_pd['Examined Entries'].fillna(0)
    exporter_pd['Examined Entries'] = exporter_pd['Examined Entries'].astype('int64')
    exporter_pd['Exam Ratio'] = (exporter_pd['Examined Entries']/exporter_pd['Total Entries'])
    exporter_pd['Exam Ratio'] = exporter_pd['Exam Ratio'].map('{:.2%}'.format)
    exporter_pd = exporter_pd[exporter_pd['Examined Entries'] > 0]

    #sheet 15: Average duties on the exam shipments, Average total invoice value on the exam shipments
    filtered = data1[((data1['Manifest Hold Crt.'].isnull() == False) & (data1['Manifest Hold Rmv.'].isnull()==False)) | ((data1['Manifest Hold Crt.'].isnull() == True) & (data1['Manifest Hold Rmv.'].isnull()==True) & (data1['Intensive Exam Opened'].isnull()==True) & (data1['CBP Hold and Exam Handling'].isnull()==False))].drop_duplicates('Entry Number')
    avg_duties = sum(filtered['Total Duty'])/len(filtered['Total Duty'])
    avg_value = sum(filtered['Total Entered Value'])/len(filtered['Total Entered Value'])
    report15 = pd.DataFrame()
    report15['Averages'] = ['Average Duties on Exam Shipements', 'Average Total Invoice Value on Exam Shipments']
    report15['Calculate Values'] = [avg_duties,avg_value]

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
    report12.to_excel(writer, index = False, sheet_name='Sheet3', startrow=1)
    sheet3 = writer.sheets['Sheet3']
    sheet3.write_string(0,0, 'All Exam in Progress Entries')
    report13.to_excel(writer, index=False, sheet_name='Sheet3', startcol=12, startrow=1)
    sheet3.write_string(0,12, 'Exam Ratios by past 90/60/30 days')
    description_pd.to_excel(writer, index = False, sheet_name='Sheet3', startcol=12, startrow=7)
    sheet3.write_string(6,12, 'Exam Ratio by Product Description (Excluding all 9903 HTS)')
    tariff_pd.to_excel(writer, index = False, sheet_name='Sheet3', startcol=17, startrow=1)
    sheet3.write_string(0,17, 'Exam Ratio by HTS')
    importer_pd.to_excel(writer, index = False, sheet_name='Sheet3', startcol=22, startrow=1)
    sheet3.write_string(0,22, 'Exam Ratio by Importer')
    exporter_pd.to_excel(writer, index = False, sheet_name='Sheet3', startcol=27, startrow=1)
    sheet3.write_string(0,27, 'Exam Ratio by Exporter')
    report15.to_excel(writer, index = False, sheet_name='Sheet3', startcol=32, startrow=1)
    sheet3.write_string(0,32, 'Average Duties on Exam Shipments')
    writer.save()

year_to_date_report(r"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\CTS_STATUS\CTS_STATUS.xlsx", r"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\CTS_STATUS\CTS_STATUS_output.xlsx")

#"C:\Users\jason\OneDrive - JH CHB\Automation\CTS_STATUS\CTS_STATUS.xlsx"
#"C:\Users\jason\OneDrive - JH CHB\Automation\CTS_STATUS\CTS_STATUS_output.xlsx"
#"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\CTS_STATUS\CTS_STATUS.xlsx"
#"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\CTS_STATUS\CTS_STATUS_output.xlsx"