import pandas as pd
import datetime

xlsx_report = pd.read_excel(r"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\ETR\ETR.xlsx")
csv_report = xlsx_report.to_csv(r"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\ETR\ETR.csv", index=False, encoding='utf-8')
file = open(r"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\ETR\ETR.csv", encoding='utf-8')

#part 1: reading the data and dropping rows if the entry number in the row is empty cell
entry = pd.read_csv(file)
entry = entry.dropna(subset=['Entry Number'])

#part 2:find all container numbers and billing invoice no for each entry number, create dict. Then remove nans in each dict
entry2 = entry.copy()
container_dict = entry2.groupby('Entry Number')['Container No'].apply(list).to_dict()
invoice_dict = entry2.groupby('Entry Number')['Billing Invoice No'].apply(list).to_dict()
for i in container_dict:
    new_list = []
    for j in container_dict[i]:
        if pd.isnull(j) == False:
            new_list.append(j)
    container_dict[i] = new_list
for i in invoice_dict:
    new_list = []
    for j in invoice_dict[i]:
        if pd.isnull(j) == False:
            new_list.append(j)
    invoice_dict[i] = new_list

#drop original container no column, add new column "Container Nos" and append list of container nos to each corresponding entry#
entry = entry.drop(columns=['Container No'])
entry['Container Nos'] = entry['Entry Number'].apply(lambda x: container_dict[x])
entry = entry.drop(columns = ['Billing Invoice No'])
entry['Billing Invoices'] = entry['Entry Number'].apply(lambda x: invoice_dict[x])

#helper functions to determine PGA status, 3 = may proceed, 2 = under review, 1 = present, 0 = not present
def epa_check(x):
    if 'EPA' in str(x):
        if 'May Proceed' in str(x):
            return 3
        elif 'Under Review' in str(x):
            return 2
        elif 'Hold' in str(x):
            return 4 
        else:
            return 1
    else:
        return 0

def fsi_check(x):
    if 'FSI' in str(x):
        if 'May Proceed' in str(x):
            return 3
        elif 'Under Review' in str(x):
            return 2
        elif 'Hold' in str(x):
            return 4 
        else:
            return 1
    else:
        return 0

def nmf_check(x):
    if 'NMF' in str(x):
        if 'May Proceed' in str(x):
            return 3
        elif 'Under Review' in str(x):
            return 2
        elif 'Hold' in str(x):
            return 4 
        else:
            return 1
    else:
        return 0

def nht_check(x):
    if 'NHT' in str(x):
        if 'May Proceed' in str(x):
            return 3
        elif 'Under Review' in str(x):
            return 2
        elif 'Hold' in str(x):
            return 4 
        else:
            return 1
    else:
        return 0

def aph_check(x):
    if 'APH' in str(x):
        if 'May Proceed' in str(x):
            return 3
        elif 'Under Review' in str(x):
            return 2
        elif 'Hold' in str(x):
            return 4 
        else:
            return 1
    else:
        return 0

def fda_check(x):
    if 'FDA' in str(x):
        if 'May Proceed' in str(x):
            return 3
        elif 'Under Review' in str(x):
            return 2
        elif 'Hold' in str(x):
            return 4
        else:
            return 1
    else:
        return 0

def ams_check(x):
    if 'AMS' in str(x):
        if 'May Proceed' in str(x):
            return 3
        elif 'Under Review' in str(x):
            return 2
        elif 'Hold' in str(x):
            return 4 
        else:
            return 1
    else:
        return 0

def ttb_check(x):
    if 'TTB' in str(x):
        if 'May Proceed' in str(x):
            return 3
        elif 'Under Review' in str(x):
            return 2
        elif 'Hold' in str(x):
            return 4 
        else:
            return 1
    else:
        return 0

def omc_check(x):
    if 'OMC' in str(x):
        if 'May Proceed' in str(x):
            return 3
        elif 'Under Review' in str(x):
            return 2
        elif 'Hold' in str(x):
            return 4 
        else:
            return 1
    else:
        return 0

def fws_check(x):
    if 'FWS' in str(x):
        if 'May Proceed' in str(x):
            return 3
        elif 'Under Review' in str(x):
            return 2
        elif 'Hold' in str(x):
            return 4 
        else:
            return 1
    else:
        return 0

def cpsc_check(x):
    if 'CPSC' in str(x):
        if 'May Proceed' in str(x):
            return 3
        elif 'Under Review' in str(x):
            return 2
        elif 'Hold' in str(x):
            return 4 
        else:
            return 1
    else:
        return 0

def lacey_check(x):
    if 'Lacey Act' in str(x):
        if 'May Proceed' in str(x):
            return 3
        elif 'Under Review' in str(x):
            return 2
        elif 'Hold' in str(x):
            return 4 
        else:
            return 1
    else:
        return 0

#part 3:make a copy of dataset, then check if each pga value is present per row by applying helper functions above
entry1 = entry.copy()
entry1['EPA'] = entry1['PGA Status'].apply(lambda x: epa_check(x))
entry1['FSI'] = entry1['PGA Status'].apply(lambda x: fsi_check(x))
entry1['NMF'] = entry1['PGA Status'].apply(lambda x: nmf_check(x))
entry1['NHT'] = entry1['PGA Status'].apply(lambda x: nht_check(x))
entry1['APH'] = entry1['PGA Status'].apply(lambda x: aph_check(x))
entry1['FDA'] = entry1['PGA Status'].apply(lambda x: fda_check(x))
entry1['AMS'] = entry1['PGA Status'].apply(lambda x: ams_check(x))
entry1['TTB'] = entry1['PGA Status'].apply(lambda x: ttb_check(x))
entry1['OMC'] = entry1['PGA Status'].apply(lambda x: omc_check(x))
entry1['FWS'] = entry1['PGA Status'].apply(lambda x: fws_check(x))
entry1['CPSC'] = entry1['PGA Status'].apply(lambda x: cpsc_check(x))
entry1['Lacey Act'] = entry1['PGA Status'].apply(lambda x: lacey_check(x))

#create dict to count all pga values per entry number
epa_dict = entry1.groupby('Entry Number')['EPA'].apply(list).to_dict()
fsi_dict = entry1.groupby('Entry Number')['FSI'].apply(list).to_dict()
nmf_dict = entry1.groupby('Entry Number')['NMF'].apply(list).to_dict()
nht_dict = entry1.groupby('Entry Number')['NHT'].apply(list).to_dict()
aph_dict = entry1.groupby('Entry Number')['APH'].apply(list).to_dict()
fda_dict = entry1.groupby('Entry Number')['FDA'].apply(list).to_dict()
ams_dict = entry1.groupby('Entry Number')['AMS'].apply(list).to_dict()
ttb_dict = entry1.groupby('Entry Number')['TTB'].apply(list).to_dict()
omc_dict = entry1.groupby('Entry Number')['OMC'].apply(list).to_dict()
fws_dict = entry1.groupby('Entry Number')['FWS'].apply(list).to_dict()
cpsc_dict = entry1.groupby('Entry Number')['CPSC'].apply(list).to_dict()
lacey_dict = entry1.groupby('Entry Number')['Lacey Act'].apply(list).to_dict()

#helper function to check the status of each entry per PGA based on the max of multiple rows, then convert back to pga status
def epa_return(x):
    if max(epa_dict[x]) == 3:
        return 'May Proceed'
    elif max(epa_dict[x]) == 2:
        return 'Under Review'
    elif max(epa_dict[x]) == 1:
        return 'No Status'
    elif max(epa_dict[x]) == 4:
        return 'Hold Intact'
    else:
        return '-'

def fsi_return(x):
    if max(fsi_dict[x]) == 3:
        return 'May Proceed'
    elif max(fsi_dict[x]) == 2:
        return 'Under Review'
    elif max(fsi_dict[x]) == 1:
        return 'No Status'
    elif max(fsi_dict[x]) == 4:
        return 'Hold Intact'
    else:
        return '-'

def nmf_return(x):
    if max(nmf_dict[x]) == 3:
        return 'May Proceed'
    elif max(nmf_dict[x]) == 2:
        return 'Under Review'
    elif max(nmf_dict[x]) == 1:
        return 'No Status'
    elif max(nmf_dict[x]) == 4:
        return 'Hold Intact'
    else:
        return '-'

def nht_return(x):
    if max(nht_dict[x]) == 3:
        return 'May Proceed'
    elif max(nht_dict[x]) == 2:
        return 'Under Review'
    elif max(nht_dict[x]) == 1:
        return 'No Status'
    elif max(nht_dict[x]) == 4:
        return 'Hold Intact'
    else:
        return '-'

def aph_return(x):
    if max(aph_dict[x]) == 3:
        return 'May Proceed'
    elif max(aph_dict[x]) == 2:
        return 'Under Review'
    elif max(aph_dict[x]) == 1:
        return 'No Status'
    elif max(aph_dict[x]) == 4:
        return 'Hold Intact'
    else:
        return '-'

def fda_return(x):
    if max(fda_dict[x]) == 3:
        return 'May Proceed'
    elif max(fda_dict[x]) == 2:
        return 'Under Review'
    elif max(fda_dict[x]) == 1:
        return 'No Status'
    elif max(fda_dict[x]) == 4:
        return 'Hold Intact'
    else:
        return '-'

def ams_return(x):
    if max(ams_dict[x]) == 3:
        return 'May Proceed'
    elif max(ams_dict[x]) == 2:
        return 'Under Review'
    elif max(ams_dict[x]) == 1:
        return 'No Status'
    elif max(ams_dict[x]) == 4:
        return 'Hold Intact'
    else:
        return '-'

def ttb_return(x):
    if max(ttb_dict[x]) == 3:
        return 'May Proceed'
    elif max(ttb_dict[x]) == 2:
        return 'Under Review'
    elif max(ttb_dict[x]) == 1:
        return 'No Status'
    elif max(ttb_dict[x]) == 4:
        return 'Hold Intact'
    else:
        return '-'

def omc_return(x):
    if max(omc_dict[x]) == 3:
        return 'May Proceed'
    elif max(omc_dict[x]) == 2:
        return 'Under Review'
    elif max(omc_dict[x]) == 1:
        return 'No Status'
    elif max(omc_dict[x]) == 4:
        return 'Hold Intact'
    else:
        return '-'

def fws_return(x):
    if max(fws_dict[x]) == 3:
        return 'May Proceed'
    elif max(fws_dict[x]) == 2:
        return 'Under Review'
    elif max(fws_dict[x]) == 1:
        return 'No Status'
    elif max(fws_dict[x]) == 4:
        return 'Hold Intact'
    else:
        return '-'

def cpsc_return(x):
    if max(cpsc_dict[x]) == 3:
        return 'May Proceed'
    elif max(cpsc_dict[x]) == 2:
        return 'Under Review'
    elif max(cpsc_dict[x]) == 1:
        return 'No Status'
    elif max(cpsc_dict[x]) == 4:
        return 'Hold Intact'
    else:
        return '-'

def lacey_return(x):
    if max(lacey_dict[x]) == 3:
        return 'May Proceed'
    elif max(lacey_dict[x]) == 2:
        return 'Under Review'
    elif max(lacey_dict[x]) == 1:
        return 'No Status'
    elif max(lacey_dict[x]) == 4:
        return 'Hold Intact'
    else:
        return '-'

#apply convert function back to dataset, create a column for each PGA status and fill each cell under each column by applying the converting helper functions above
entry['EPA'] = entry['Entry Number'].apply(lambda x: epa_return(x))
entry['FSI'] = entry['Entry Number'].apply(lambda x: fsi_return(x))
entry['NMF'] = entry['Entry Number'].apply(lambda x: nmf_return(x))
entry['NHT'] = entry['Entry Number'].apply(lambda x: nht_return(x))
entry['APH'] = entry['Entry Number'].apply(lambda x: aph_return(x))
entry['FDA'] = entry['Entry Number'].apply(lambda x: fda_return(x))
entry['AMS'] = entry['Entry Number'].apply(lambda x: ams_return(x))
entry['TTB'] = entry['Entry Number'].apply(lambda x: ttb_return(x))
entry['OMC'] = entry['Entry Number'].apply(lambda x: omc_return(x))
entry['FWS'] = entry['Entry Number'].apply(lambda x: fws_return(x))
entry['CPSC'] = entry['Entry Number'].apply(lambda x: cpsc_return(x))
entry['Lacey Act'] = entry['Entry Number'].apply(lambda x: lacey_return(x))

#date difference
current_date = datetime.date.today()
entry['Date Difference'] = pd.to_datetime(entry['Arrival Date'])
entry['Date Difference'] = entry['Date Difference'].dt.date
entry['Date Difference'] = entry['Date Difference'] - current_date


#convert date time format 
entry['Entry Date'] = pd.to_datetime(entry['Entry Date'], format='%Y-%m-%d').dt.strftime('%m/%d/%Y')
entry['Import Date'] = pd.to_datetime(entry['Import Date'], format='%Y-%m-%d').dt.strftime('%m/%d/%Y')
entry['Arrival Date'] = pd.to_datetime(entry['Arrival Date'], format='%Y-%m-%d').dt.strftime('%m/%d/%Y')
entry['Statement Date'] = pd.to_datetime(entry['Statement Date'], format='%Y-%m-%d').dt.strftime('%m/%d/%Y')
entry['Cargo Release Date'] = pd.to_datetime(entry['Cargo Release Date'], format='%Y-%m-%d').dt.strftime('%m/%d/%Y')
entry['Intensive Exam Compl.'] = pd.to_datetime(entry['Intensive Exam Compl.'], format='%Y-%m-%d').dt.strftime('%m/%d/%Y')
#entry['Date'] = pd.to_datetime(entry['Date'], format='%Y-%m-%d').dt.strftime('%m/%d/%Y')
entry['Billing Invoice Date'] = pd.to_datetime(entry['Billing Invoice Date'], format='%Y-%m-%d').dt.strftime('%m/%d/%Y')





#drop all duplicate entry number rows
entry = entry.drop_duplicates(subset=['Entry Number'])
entry = entry.reset_index(drop=True)

#write resulting dataset into csv
#entry.to_excel(r"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\ETR\ETR_output.xlsx", index=False)
writer = pd.ExcelWriter(r"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\ETR\ETR_output.xlsx", engine='xlsxwriter') 
entry.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)

workbook = writer.book
worksheet = writer.sheets['Sheet1']

(max_row, max_col) = entry.shape

column_settings = [{'header': column} for column in entry.columns]

worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

worksheet.set_column(0, max_col - 1, 12)

writer.save()
