import os
import pandas as pd
import xlwings as xw

folder_path = r"T:\Project\aeT00989.00 Decarbonization Feasibility Study\03. Client design information\3.01 General\Needlist-4\b1"

Master_Index_path=r"T:\Project\aeT00989.00 Decarbonization Feasibility Study\06. Tebodin Design Information (No products)\6.01 General\Data Sorted\void\Master Index.xlsx"

Master_Index=pd.read_excel(Master_Index_path,header=None)
df_combined=Master_Index.iloc[:,:9]
cols=df_combined.iloc[0]
df_combined=df_combined[1:]
df_combined.columns=cols

ErrorList = []

for root, dirs, files in os.walk(folder_path):
    for file in files:
        # try:
        if file.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(os.path.join(root, file), header=None, skiprows=4,index_col=None)
            df=df.drop(df.columns[5],axis=1)
            # df = df.reindex(columns=df_combined.columns)
            df['Need List Name'] = file
            
            
            if not df['Need List Name'].isin(df_combined["Need List Name"]).any():
                df["NL Batch"]=folder_path.split("\\")[-1]
                df.columns=cols
                df_combined=pd.concat([df_combined,df],axis=0)

df_combined=df_combined.dropna(subset=["Discipline","Doc. No."])

df_combined = df_combined[~(df_combined['Discipline'].str.contains('Discipline'))]
df_combined =df_combined.reset_index(drop=True)
df_combined = df_combined.drop("S.No.",axis=1)
df_combined.index.name="S.No."

app = xw.App(visible=False)
excel_file = Master_Index_path
book = xw.Book(excel_file)
sheet_name = 'Sheet1'
sheet = book.sheets[sheet_name]
sheet.range('A2').options(index=True, header=False).value = df_combined

book.save()
book.close()
app.quit()

