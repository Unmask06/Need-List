import pandas as pd
import sqlite3
from datetime import datetime
import traceback
import xlwings as xw

# Paths
excel_file_path = r"T:\Project\aeT00989.00 Decarbonization Feasibility Study\06. Tebodin Design Information (No products)\6.01 General\Data Sorted\void\Master Index.xlsx"
sheet_name = "Unmapped"
database_path = "database.db"
db_table = "tbl_ADNOC"
columns_to_fill = ["Description", "Discipline", "Site"]


# Connect to the database
def load_sql_data(database_path, db_table):
    conn = sqlite3.connect(database_path)
    c = conn.cursor()

    query = f"SELECT * FROM '{db_table}'"
    sql_df = pd.read_sql_query(query, conn)
    conn.commit()
    conn.close()

    return sql_df


def read_and_clean_excel_data(excel_file_path):
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=0).fillna("")
    df.drop(df.columns[[0, 1]], axis=1, inplace=True)
    df["Doc. No."].fillna("Doc. Number missing", inplace=True)
    global selected_columns
    selected_columns = df.columns.values.tolist()
    return df


def fill_data(df, sql_df, columns_to_fill):
    for i, row in df.iterrows():
        try:
            if any(row[col] == "" or pd.isna(row[col]) for col in columns_to_fill):
                document = str(row["Doc. No."]).replace("/", "")
                matching_row = sql_df.loc[sql_df["Document No."].str.replace("/", "") == document]
                if not matching_row.empty:
                    description = matching_row["Description / Title"].values[0]
                    discipline = matching_row["Discipline"].values[0]
                    site = matching_row["Site"].values[0]
                    # sector = matching_row['Sector'].values[0]

                    if df.at[i, "Description"] == "" or pd.isna(row["Description"]):
                        df.at[i, "Description"] = description
                    if df.at[i, "Discipline"] == "" or pd.isna(row["Discipline"]):
                        df.at[i, "Discipline"] = discipline
                    if df.at[i, "Site"] == "" or pd.isna(row["Site"]):
                        df.at[i, "Site"] = site
                    # if pd.isna(df.at[i, 'Sector']):
                    #     df.at[i, 'Sector'] = sector

        except Exception as e:
            print(f"Error occurred at index {i}: {traceback.format_exc()}")

    return df


def write_to_excel(df, excel_file_path, sheet_name):
    app = xw.App(visible=False)
    wb = xw.Book(excel_file_path)
    UnNamed_sh = wb.sheets[sheet_name]
    UnNamed_sh.range("B2").options(index=True, header=False).value = df
    wb.save(excel_file_path)
    wb.close()
    app.quit()


sql_df = load_sql_data(database_path, db_table)

df = read_and_clean_excel_data(excel_file_path)

Filled_df = fill_data(df, sql_df, columns_to_fill)

write_to_excel(Filled_df, excel_file_path, sheet_name)
