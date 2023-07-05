
import pandas as pd
import sqlite3
import xlwings as xw
import traceback

excel_file_path = r"T:\Project\aeT00989.00 Decarbonization Feasibility Study\06. Tebodin Design Information (No products)\6.01 General\Data Sorted\void\Master Index.xlsx"
sheet_name="Sheet1"
conn = sqlite3.connect('database.db')
c=conn.cursor()

query="SELECT * FROM 'tbl_ADNOC'"
sql_df=pd.read_sql_query(query,conn)

conn.commit()

df=pd.read_excel(excel_file_path,sheet_name=sheet_name,header=0,index_col=1)
df=df.drop(columns=["File Link"])


for i, row in df.iterrows():
    try:
        description = row["Description"]
        discipline = row["Discipline"]
        
        if pd.isna(description):
            document = str(row['Doc. No.']).replace("/", "")
            sql_row = sql_df.loc[sql_df["Document No."].str.replace("/", "") == document]
            
            if not sql_row.empty:
                description = sql_row.iloc[0]["Description / Title"]
                if pd.isna(discipline):
                    discipline = sql_row.iloc[0]["Discipline"]
                    
                df.at[i, "Description"] = description
                df.at[i, "Discipline"] = discipline
    except:
        print(i, traceback.format_exc())



selected_columns = [
    'Discipline', 'Description', 'Doc. No.', 'Rev', 'Status', 'Remarks',
    'Need List Name', 'NL Batch', 'Site', 'Type', 'File Path', 'File count', 'Processed Date','Source Path'
]

df_selected=df[selected_columns]

app = xw.App(visible=False)
wb =xw.Book(excel_file_path)
sheet = wb.sheets[0]
UnNamed_sh=wb.sheets[sheet_name]

UnNamed_sh.range("B2").options(index=True,header=False).value=df_selected

wb.save(excel_file_path)
wb.close()
app.quit()





