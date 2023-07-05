import pandas as pd
import os
import xlwings as xw
from datetime import datetime
import traceback

excel_file_path = r"T:\Project\aeT00989.00 Decarbonization Feasibility Study\06. Tebodin Design Information (No products)\6.01 General\Data Sorted\void\Master Index.xlsx"


master_folder_path = r'T:\Project\aeT00989.00 Decarbonization Feasibility Study\03. Client design information\3.01 General\As Received\09-05-2023_UZ-Requested Data'
file_list=[]
file_list_links=[]

for root, dirs, files in os.walk(master_folder_path):
    for file in files:
        filename,extension=os.path.splitext(file)
        if extension!=['.zip','.DWG']:
            file_list.append(filename)
            file_list_links.append(os.path.join(root,file))

selected_columns = ['Discipline', 'Description', 'Doc. No.', 'Rev', 'Status', 'Remarks',
    'Need List Name', 'NL Batch', 'Site', 'Type', 'File Path', 'File count', 'Processed Date','Source Path','File Link']

dfFile=pd.DataFrame(columns=selected_columns)
dfFile["Doc. No."]=file_list
dfFile["File Path"]=file_list_links
dfFile["Processed Date"]=datetime.today().date().strftime("%d-%b-%y")
dfFile["NL Batch"]=master_folder_path.split("\\")[-1]

app = xw.App(visible=False)
wb =xw.Book(excel_file_path)
sheet = wb.sheets["Individual Files"]

if sheet.range("B2").value!=None:
    last_row=(sheet.range("B1").end("down").row)+1
else:
    last_row=2
sheet.range(f"B{last_row}").options(index=True,header=False).value=dfFile

# for i in range(last_row,len(dfFile)+last_row):
#     sheet.range(f'P{i}').value=f'=HYPERLINK(L{i},D{i})'
#     sheet.range(f'A{i}').value=i

wb.save(excel_file_path)
wb.close()
app.quit()



        
