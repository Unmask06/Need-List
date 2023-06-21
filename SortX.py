import os
import pandas as pd
import xlwings as xw
import shutil
from datetime import datetime
import logging


class MasterIndex:
    def __init__(self, config_file_path="config.xlsx", overwrite_log=True):
        self.log_file = "sortx.log"
        self.setup_logging(overwrite_log)
        self.load_config(config_file_path)
        self.load_mapper(config_file_path)
        self.check_master_index()

    def setup_logging(self, overwrite_log=True):
        logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

        if overwrite_log:
            with open(self.log_file, "w"):
                pass
            root_logger = logging.getLogger()
            for handler in root_logger.handlers:
                root_logger.removeHandler(handler)

        file_handler = logging.FileHandler(self.log_file)
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        logging.getLogger().addHandler(file_handler)

    def load_config(self, config_file_path):
        dfconfig = pd.read_excel(config_file_path, sheet_name="config", header=0)
        self.config = dict(zip(dfconfig.iloc[:, 0], dfconfig.iloc[:, 1]))

    def load_mapper(self, config_file_path):
        dfmapper = pd.read_excel(config_file_path, sheet_name="mapper", header=0)
        self.mapper = dict(zip(dfmapper.iloc[:, 0], dfmapper.iloc[:, 1]))
        self.mandate_columns = list(self.mapper.values())

    def check_master_index(self):
        self.path = self.config["master_index_path"]
        if not os.path.exists(self.path):
            error_msg = f"Master index file not found at {self.path}"
            logging.error(error_msg)
            raise FileNotFoundError(error_msg)
        else:
            logging.info(f"Master index file found at {self.path}")

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

        except (ValueError, FileNotFoundError) as e:
            logging.error(e)
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

folder_path = r"Need Lists\NL 1\ELECTRICA"

dfmerged = mi.merge_excel(folder_path)

# mi.write_to_excel(dfmerged)
