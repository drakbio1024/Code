import os
import glob

list_of_downloads = glob.glob('C:\\Users\\q1071018\\Downloads\\*.xlsx')

open_list = []
full_list = []
cons_list = []

for item in list_of_downloads :
    if item.startswith('C:\\Users\\q1071018\\Downloads\\Q2_Order_Report'):
        full_list.append(item)
    if item.startswith('C:\\Users\\q1071018\\Downloads\\Q2_Consolidated_Order'):
        cons_list.append(item)
    if item.startswith('C:\\Users\\q1071018\\Downloads\\Q2_Open_Order_Report'):
        open_list.append(item)

if len(open_list) != 0:
    latestopen = max(open_list, key=os.path.getctime)
    for item in open_list:
        if item != latestopen:
            os.remove(item)
            print(item + ' deleted.') 


if len(full_list) != 0:
    latestfull = max(full_list, key=os.path.getctime)
    for item in full_list:
        if item != latestfull:
            os.remove(item)  
            print(item + ' deleted.') 

if len(cons_list) != 0:      
    for item in cons_list:
        os.remove(item) 
        print(item + ' deleted.') 

print('Finished.')  
