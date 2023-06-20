import os
import pandas as pd
import xlwings as xw
import shutil
from datetime import datetime


class MasterIndex:
    def __init__(self, config_file_path="config.xlsx"):
        dfconfig = pd.read_excel(config_file_path, sheet_name="config", header=0)
        self.config = dict(zip(dfconfig.iloc[:, 0], dfconfig.iloc[:, 1]))

        dfmapper = pd.read_excel(config_file_path, sheet_name="mapper", header=0)
        self.mapper = dict(zip(dfmapper.iloc[:, 0], dfmapper.iloc[:, 1]))

        self.mandate_columns = list(self.mapper.values())

        self.path = self.config["master_index_path"]
        if not os.path.exists(self.path):
            raise FileNotFoundError(f"Master index file not found at {self.path}")

    def merge_excel(self, folder_path: str, index_col=0) -> pd.DataFrame:
        try:
            dfs = []
            skip_rows = self.config["header_row_number"] - 1
            index_col = self.config["sno_column"] - 1 if self.config["sno_column"] != "" else 0

            for file in os.listdir(folder_path):
                if file.endswith((".xlsx", ".xls")):
                    df = pd.read_excel(
                        os.path.join(folder_path, file), skiprows=skip_rows, index_col=index_col
                    )
                    df = df[self.mandate_columns]
                    dfs.append(df)
            dfmerged = pd.concat(dfs, ignore_index=True)
            reversed_mapper = {v: k for k, v in self.mapper.items()}

            # TODO: Add a check to see if all column names are same across all files

            dfmerged = dfmerged.rename(columns=reversed_mapper)
            return dfmerged
        except ValueError as e:
            print(e)

    def write_to_excel(self, df, sheet_name=0):
        app = xw.App(visible=False)
        excel_file = self.path
        book = xw.Book(excel_file)
        sheet = book.sheets[sheet_name]

        # TODO: Add a variable to ask user if he wants to append or overwrite

        last_row = sheet.api.Cells(sheet.api.Rows.Count, "B").End(-4162).Row
        sheet.range(f"B{last_row+1}").options(index=True, header=True).value = df
        book.save()
        book.close()
        app.quit()


mi = MasterIndex()

folder_path = r"Need Lists\NL 1\ELECTRICAL"

dfmerged = mi.merge_excel(folder_path)

mi.write_to_excel(dfmerged)
