import os
import pandas as pd
import shutil
import traceback
from datetime import datetime
import xlwings as xw

excel_file_path = r"T:\Project\aeT00989.00 Decarbonization Feasibility Study\06. Tebodin Design Information (No products)\6.01 General\Data Sorted\void\Master Index.xlsx"
master_folder_path = r"T:\Project\aeT00989.00 Decarbonization Feasibility Study\03. Client design information\3.01 General\As Received\29-05-2023_UL & US"
new_folder_path = r"T:\Project\aeT00989.00 Decarbonization Feasibility Study\06. Tebodin Design Information (No products)\6.01 General\Data Sorted"


def read_and_clean_excel_data(file_path):
    df = pd.read_excel(file_path, index_col=1)
    df.drop(df.columns[0], axis=1, inplace=True)
    df["Doc. No."].fillna("Doc. Number missing", inplace=True)
    global selected_columns
    selected_columns = df.columns.values.tolist()
    return df


def process_data(df, master_folder_path, new_folder_path):
    for root, dirs, files in os.walk(master_folder_path):
        for dir_item in dirs:
            doc = dir_item
            for index in range(len(df)):
                try:
                    process_item(df, index, doc, root, dir_item, new_folder_path)
                except Exception as e:
                    handle_error(df, index)
    return df[selected_columns]


def process_item(df, index, doc, root, dir_item, new_folder_path):
    if df.at[index, "Status"] != "closed":
        if str(df.at[index, "Doc. No."]) in doc:
            df.at[index, "Status"] = "closed"
            df.at[index, "Processed Date"] = datetime.today().date().strftime("%d-%b-%y")

            set_source_and_destination_path(df, index, root, dir_item, new_folder_path)
            copy_files(df, index)


def set_source_and_destination_path(df, index, root, dir_item, new_folder_path):
    item_path = os.path.join(root, dir_item)
    df.at[index, "Source Path"] = item_path

    discipline_folder = os.path.join(new_folder_path, df.at[index, "Discipline"])
    site_folder = os.path.join(discipline_folder, df.at[index, "Site"])

    doc_type = df.iloc[index]["Type"] if not pd.isnull(df.iloc[index]["Type"]) else "Uncategorized"
    doc_type_folder = os.path.join(site_folder, doc_type)

    destination_path = os.path.join(doc_type_folder, dir_item)
    df.at[index, "File Path"] = destination_path

    os.makedirs(site_folder, exist_ok=True)
    os.makedirs(doc_type_folder, exist_ok=True)


def copy_files(df, index):
    if not os.path.exists(df.at[index, "File Path"]):
        shutil.copytree(str(df.at[index, "Source Path"]), str(df.at[index, "File Path"]))


def handle_error(df, index):
    df.at[index, "Status"] = ""
    df.at[index, "Processed Date"] = ""
    print(index, traceback.format_exc())


def create_folder_data(master_folder_path, excel_file_path):
    last_level_folders = []
    Folders_link = []

    for root, dirs, files in os.walk(master_folder_path):
        if not dirs:
            Folders_link.append(root)
            last_folder_name = os.path.basename(root)
            last_level_folders.append(last_folder_name)
    File_df = pd.read_excel(excel_file_path, sheet_name="Unmapped")
    File_df = File_df.drop(index=File_df.index)
    File_df = File_df[selected_columns]
    File_df = File_df.reset_index(drop=True)
    File_df["Doc. No."] = last_level_folders
    File_df["File Path"] = Folders_link
    File_df["Mapped"] = ""
    return File_df


def map_folders(df, File_df):
    basenames = [os.path.basename(path) for path in df["File Path"] if isinstance(path, str)]
    for i in range(len(File_df)):
        folder_name = File_df["Doc. No."][i]
        File_df.at[i, "Processed Date"] = datetime.today().date().strftime("%d-%b-%y")
        if folder_name in basenames:
            index = basenames.index(folder_name)
            File_df.loc[i, "Mapped"] = True
            File_df.loc[i, "Discipline"] = df.at[index, "Discipline"]
        else:
            File_df.loc[i, "Mapped"] = False
    UnNamed = File_df.loc[File_df["Mapped"] == False]
    UnNamed = UnNamed[selected_columns]
    return UnNamed


def write_to_excel(df_selected, UnNamed, excel_file_path):
    app = xw.App(visible=False)
    wb = xw.Book(excel_file_path)
    sheet = wb.sheets[0]
    UnNamed_sh = wb.sheets["Unmapped"]
    sheet.range("B2").options(index=True, header=False).value = df_selected
    if UnNamed_sh.range("B2").value != None:
        last_row = (UnNamed_sh.range("B1").end("down").row) + 1
    else:
        last_row = 2
    UnNamed_sh.range(f"B{last_row}").options(index=True, header=False).value = UnNamed
    wb.save(excel_file_path)
    wb.close()
    app.quit()


df = read_and_clean_excel_data(excel_file_path)
df_selected = process_data(df, master_folder_path, new_folder_path)
File_df = create_folder_data(master_folder_path, excel_file_path)
UnNamed = map_folders(df, File_df)
write_to_excel(df_selected, UnNamed, excel_file_path)
