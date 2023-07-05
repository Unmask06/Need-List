import pandas as pd
import os
import xlwings as xw
import shutil
from datetime import datetime
import traceback

excel_file_path = r"T:\Project\aeT00989.00 Decarbonization Feasibility Study\06. Tebodin Design Information (No products)\6.01 General\Data Sorted\void\Master Index.xlsx"
df = pd.read_excel(excel_file_path,sheet_name="Unmapped",index_col=1)
df=df.drop(columns=["File Link"])

df=df.dropna(subset=["Discipline"])
df.reset_index(drop=True,inplace=True)
df["File Path"], df["Error"] = "", ""

new_folder_path = r"T:\Project\aeT00989.00 Decarbonization Feasibility Study\06. Tebodin Design Information (No products)\6.01 General\Data Sorted"

for index in range(len(df)):
    try:
        if df.at[index, "Status"] != "closed":
            if str(df.at[index, "Doc. No."]) in df.at[index, "File Path"]:
                df.at[index, "Status"] = "closed"
                df.at[index, "Processed Date"] = datetime.today().date().strftime("%d-%b-%y")

                file_name=str(os.path.basename(df.at[index, "File Path"]))
                df.at[index, "Source Path"] = df.at[index, "File Path"]

                discipline_folder = os.path.join(new_folder_path, df.at[index, "Discipline"])
                site_folder = os.path.join(discipline_folder, df.at[index,"Site"])

                doc_type = df.iloc[index]["Type"] if not pd.isnull(df.iloc[index]["Type"]) else "Uncategorized"
                doc_type_folder = os.path.join(site_folder, doc_type)

                destination_path = os.path.join(doc_type_folder, file_name)
                df.at[index, "File Path"] = destination_path

                os.makedirs(site_folder, exist_ok=True)
                os.makedirs(doc_type_folder, exist_ok=True)

                if not os.path.exists(destination_path):
                    shutil.copytree(df.iloc[index]["Source Path"], df.iloc[index]["File Path"])

                df.at[index, "File Path"] = df.at[index, "File Path"]

            else:
                continue
        else:
            continue
    except Exception as e:
        df.at[index, "Status"] =""
        print(traceback.format_exc())

df=df.iloc[:,:14]

app = xw.App(visible=False)
wb =xw.Book(excel_file_path)
sheet = wb.sheets[0]

if sheet.range("B2").value!=None:
    last_row=(sheet.range("B1").end("down").row)+1
else:
    last_row=2
sheet.range(f"B{last_row}").options(index=True,header=False).value=df

# for i in range(last_row,len(df)+last_row):
#     sheet.range(f'P{i}').value=f'=IF(ISBLANK(L{i}),"",HYPERLINK(L{i},D{i}))'
#     sheet.range(f'A{i}').value=i-2

wb.save(excel_file_path)
wb.close()
app.quit()