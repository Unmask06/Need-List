#Site Segregation
# sites={"UZ":["Upper Zakum","UZ"],
#        "LZ":["Lower Zakum","LZ"],
#        "UL":["UL","Umm-Lulu"],
#        "NASR":["NASR"],
#        "UAD":["UAD"],
#        "SATAH":["SATAH"],
#        "DAS":["DAS"],
#        "ZIRKU":["ZIRKU"],
#        "ABK":["ABK"],
#        "AR":["ARZAHNA","AR"],
#        "US":["UMM SHAIF","US"]}

sites={"UZ":["Upper Zakum"],
       "LZ":["Lower Zakum"],
       "UL":["Umm-Lulu"],
       "NASR":["NASR"],
       "UAD":["UAD"],
       "SATAH":["SATAH"],
       "DAS":["DAS"],
       "ZIRKU":["ZIRKU"],
       "ABK":["ABK"],
       "AR":["ARZAHNA"],
       "US":["UMM SHAIF"]}

doc_types={"PFD":["PFD","FLOW DIAG"],
          "UFD":["UFD","Utility Flow Dia"],
          "PSD":["PSD","Safety Diagram","PROCESS SAFEGUARDING FLOW","PROCESS SAFETY DIAGRAM"],
          "P&ID":["PID","P&ID","PIPING AND INSTRUMENT","PIPING & INSTRUMENT","P and ID"],
          "SLD":["SLD","Single Line Dia","Line Dia"],
          "Key Diagram":["Key Dia"],
          "Layout":["Layout"],
          "Equipment List":["Equipment List"],
          "HMB":["HMB","Heat and Material","Material balance"],
          "Load List":["Load List"],
          "Plot Plan":["Plot Plan"],
          "Data Sheet":["Datasheet","Datasht"],
          "Loop Drawing":["LOOP DRAWING"," LOOP DIAGRAM"],
          "Routing":["CABLE ROUTING"],
          "Wiring Diagram" : ["WIRING DIAGRAM"],
          "Area Classification" : ["AREA CLASSIFICATIONS","HAZARDAUS AREA","HAZARDOUS AREA"],
          "IO List" : ["INPUT/OUTPUT LIST","I/O LIST"],
          "C&E Diagram" : ["CAUSE & EFFECT","CAUSE AND EFFECT"],
          "ESD" : ["EMERGENCY SHUTDOWN DIAGRAM"],
          "Piping Plan" : ["PIPING PLAN"],
          "Assembly Drawing" : ["ASSEMBLY DWG"],
          "Calculation" : ["CALCULATION"],
          "DOSSIER" : ["DOSSIER"],
          "Philosophy" : ["Philosophy"],
          "Key Plan" : ["KEY PLAN"],
          "JUNCTION BOX" : ["JUNCTION BOX"]
            }
depts={"PROCESS":["PROCESS"],
      "ELECTRICAL":["ELECTRICAL"],
      "CIVIL & STRUCTURAL":["Civil","STRUCTURAL","Structures"],
    "INSTRUMENTATION":["INSTRUMENTATION","TELECOMMUNICATION"],
    "General":["General"],
    "MECHANICAL":["MECH","Piping","Equipments"],
    "HVAC":["HVAC"],
    "SAFETY":["LOSS PREVENTION","Safety"]
}



import pandas as pd
import xlwings as xw
import traceback

def load_excel_file(file_path, sheet_name,omit_col):
    """Load excel file into pandas dataframe."""
    df = pd.read_excel(file_path, sheet_name=sheet_name).fillna("")
    df = df.drop(df.columns[omit_col],axis=1)
    return df

def update_columns(df, mapping_dict, column_to_update):
    """Update dataframe columns based on mapping dictionary."""
    for key, values in mapping_dict.items():
        values = ("|".join(values)).upper().replace(" ", "")
        conditions = [((df[col].astype(str).str.upper()).str.replace(" ","")).str.contains(values) for col in df.columns]
        bool_series = pd.concat(conditions, axis=1).any(axis=1) & df[column_to_update].str.contains("")
        df.loc[bool_series, [column_to_update]] = key
    return df

def update_discipline(df, depts):

    for key, values in depts.items():
        values = ("|".join(values)).upper().replace(" ", "")
        conditions = ((df["Discipline"].str.upper()).str.replace(" ","")).str.contains(values)
        bool_series = conditions
        df.loc[bool_series, "Discipline"] = key
    return df

def save_excel_file(df, file_path, sheet_name):
    """Save dataframe to excel file."""
    app = xw.App(visible=False)
    book = xw.Book(file_path)
    sheet = book.sheets[sheet_name]
    sheet.range('B2').options(index=True, header=False).value = df
    book.save()
    book.close()
    app.quit()

def main(file_path,sheet_name,df,columns_to_update,mapping_dicts):
    """Main function."""
   
    
    for mapping_dict, column_to_update in zip(mapping_dicts, columns_to_update):
        df = update_columns(df, mapping_dict, column_to_update)
        
    df = update_discipline(df,depts)

    save_excel_file(df, file_path, sheet_name)
    

file_path = r"T:\Project\aeT00989.00 Decarbonization Feasibility Study\06. Tebodin Design Information (No products)\6.01 General\Data Sorted\void\Master Index.xlsx"
sheet_name = "Unmapped"
df = load_excel_file(file_path, sheet_name,omit_col=[0,1])
columns_to_update = ['Discipline', 'Site', 'Type']
mapping_dicts=[depts,sites,doc_types]

if __name__ == "__main__":
    main(file_path,sheet_name,df,columns_to_update,mapping_dicts)
