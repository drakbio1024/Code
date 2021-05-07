import sys, glob, os
from winreg import *

import openpyxl as xl  
import pandas as pd
import datetime as dt
import styleframe as sf


print('make sure you have saved TOPCAT GREEN LIGHT MASTER FILE and Pending approvals to the Downloads/CI downloads folder.')
#input('press enter to continue...')

#create DataFrame out of the 2 CI approval files, 2 columns "Order No" and "DATE SENT TO BUILD"
path = os.path.join(sys.path[0], r"Spreadsheets\TOPCAT GREEN LIGHT MASTER FILE.xlsx")
sheet = pd.read_excel(path, sheet_name = 'WORK ORDERS')
orders1 = sheet[['Order No', 'DATE SENT TO BUILD']]

path = os.path.join(sys.path[0], r"Spreadsheets\Pending Approval.xlsx")
sheet = pd.read_excel(path, sheet_name = 'Outbound log')
orders2 = sheet[['Order Number', 'Notify Therapak on approval to ship']]
orders2.columns = ['Order No', 'DATE SENT TO BUILD']

combined = orders1.append(orders2)

#take user input, transfer format
print('Please input desired sent-to-build date in format YYYY-MM-DD:')
date = input()

datetransfer = {'01' : 'Jan', '02' : 'Feb', '03' : 'Mar', '04' : 'Apr', '05' : 'May', '06' : 'Jun', 
                '07' : 'Jul', '08' : 'Aug', '09' : 'Sep', '10' : 'Oct', '11' : 'Nov', '12' : 'Dec' }
date2 = date[-2:] + '-' + datetransfer[date[5:7]] + '-' + date[0:4]

print('Please input date to update Target Ship Dates to (YYYY-MM-DD): ')
TSDdate = input()
TSDdate_int = int(TSDdate[0:4] + TSDdate[5:7] + TSDdate[-2:])
TSDdate2 = TSDdate[-2:] + '-' + datetransfer[TSDdate[5:7]] + '-' + TSDdate[0:4]

#subset combined df using date input to DF that contains only orders with build date of inputted date
#also subset to get only TOPCAT orders. Stored in orders
DateTime = dt.datetime.strptime(date + ' 00:00:00', '%Y-%m-%d %H:%M:%S')
combined_subset_1 = combined[combined['DATE SENT TO BUILD'] == DateTime]
orders = combined_subset_1[combined_subset_1['Order No'].str.startswith('Q000')]

#find latest order report within downloads, import into df_report, set columns to pull TSD and QSQ#
with OpenKey(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
    Downloads = QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]

list_of_files = glob.glob(Downloads + r'\*')

list_of_reports = []
for item in list_of_files :
    if item.startswith(Downloads + r"\Q2_Order_Report"):
        list_of_reports.append(item)

latest_report = max(list_of_reports, key=os.path.getctime)
reportpath = latest_report

df = pd.read_excel(reportpath, sheet_name = 'Report')
df_report = df[['Order #', 'PeopleSoft Order #', 'Target Ship Date']]
df_report = df_report.set_index('PeopleSoft Order #')

#set index of orders as order no, then concatenate orders and TSDs from report
orders_indexed = orders.set_index('Order No') 
s1 = df_report['Target Ship Date']
s2 = df_report['Order #']
orders2 = pd.concat([orders_indexed, s1], axis=1)
orders3 = pd.concat([orders2, s2], axis = 1)
orders4 = orders3[['Order #', 'DATE SENT TO BUILD', 'Target Ship Date']]

#getting any rows with NaN and exporting to ordersNF, to get them later when they are inputted
ordersNF = orders4[orders4['Target Ship Date'].isnull()]

#drop all NaN from orders4, then drop the DSTB column, for export into TWIST
orders55 = orders4.dropna()
orders5 = orders55.drop('DATE SENT TO BUILD', axis = 1)

#insert other columns for TWIST export
orders5.insert(loc = 1, column = 'Append Internal Notes', value = 'CI Approved ' + date2)
orders5.insert(loc = 1, column = 'Delay Reason', value = 'Sponsor Resolved')
orders5.insert(loc = 1, column = 'Order Status', value = 'In Process')
orders5 = orders5.rename(columns = {'Order #' : 'Order No'})

#change TSD of any orders that are set to ship before the inputted date
orders5['Target Ship Date'].loc[orders5['Target Ship Date'] < TSDdate] = TSDdate


#export to spreadsheet
orders4.to_excel(os.path.join(sys.path[0], r"Not Found Orders\not uploaded orders " + date + ".xlsx"), sheet_name = date)
orders5.to_excel(os.path.join(sys.path[0], r"Spreadsheets\Todays CI uploads.xlsx"), index = False)

