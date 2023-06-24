import os
import pandas as pd
import xlwings as xw
import shutil
from datetime import datetime
import logging
import traceback


class CustomException(BaseException):
    pass


class MasterIndex:
    def __init__(self, config_file_path="config.xlsm", overwrite_log=True):
        self.log_file = "sortx.log"
        self.setup_logging(overwrite_log)
        self.load_config(config_file_path)
        self.load_mapper(config_file_path)
        self.load_required_columns(config_file_path)
        self.load_master_index()

    def setup_logging(self, overwrite_log=True):
        logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

        if overwrite_log:
            with open(self.log_file, "w"):
                pass
            self.logger = logging.getLogger(__name__)
            for handler in self.logger.handlers:
                self.logger.removeHandler(handler)

        file_handler = logging.FileHandler(self.log_file)
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        self.logger.addHandler(file_handler)

        return self.logger

    def load_config(self, config_file_path):
        dfconfig = pd.read_excel(config_file_path, sheet_name="config", header=0).fillna("")
        self.config = dict(zip(dfconfig.iloc[:, 0], dfconfig.iloc[:, 1]))

    def load_mapper(self, config_file_path):
        dfmapper = pd.read_excel(config_file_path, sheet_name="mapper", header=0)
        self.mapper = dict(zip(dfmapper.iloc[:, 0], dfmapper.iloc[:, 1]))
        self.mandate_columns = list(self.mapper.values())

    def load_required_columns(self, config_file_path):
        dfrequired = pd.read_excel(config_file_path, sheet_name="field", header=0)
        self.required_columns = list(self.mapper.keys()) + list(dfrequired.iloc[:, 0])

    def load_master_index(self):
        try:
            self.path = self.config["master_index_path"]
            if os.path.exists(self.path):
                self.logger.info(f"Master index file found at {self.path}")
            elif self.path == "":
                error_msg = "Master index path not specified in config file"
                self.logger.error(error_msg)
                raise CustomException(error_msg)

        except (FileNotFoundError, TypeError) as e:
            error_msg = f"{e}\nMaster index file not found at {self.path}"
            self.logger.error(error_msg)
            raise CustomException(error_msg)

        try:
            dfmaster = pd.read_excel(self.path, sheet_name=0, header=0)
            dfmaster = dfmaster[self.required_columns]
            if not set(self.required_columns).issubset(set(dfmaster.columns)):
                missing_columns = set(self.required_columns) - set(dfmaster.columns)
                error_msg = (
                    'Ensure Master Index has all columns specified in the config file ("A" columns of mapper + field sheet)\n'
                    f"MISSING COLUMNS: {', '.join(missing_columns)}\n"
                    "Please update those columns in the master index and try again."
                )
                self.logger.error(error_msg)
                raise ValueError(error_msg)
            else:
                self.dfmaster = dfmaster

        except Exception as e:
            error_msg = f"{e}\nError while reading master index file"
            self.logger.error(error_msg)
            raise CustomException(error_msg)

    def merge_excel(self, folder_path):
        try:
            dfs = []
            skip_rows = self.config["header_row_number"] - 1
            index_col = self.config["sno_column"] - 1 if self.config["sno_column"] != "" else 0

            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    if file.endswith((".xlsx", ".xls")):
                        df = pd.read_excel(
                            os.path.join(root, file), skiprows=skip_rows, index_col=index_col
                        )
                        df = df[self.mandate_columns]
                        df["imported_from"] = file
                        if set(self.dfmaster["imported_from"]).isdisjoint(set(df["imported_from"])):
                            dfs.append(df)
            dfmerged = pd.concat(dfs, ignore_index=True)
            reversed_mapper = {v: k for k, v in self.mapper.items()}

            # TODO: Add a check to see if all column names are same across all files

            # TODO: Add a column for excel file name

            dfmerged = dfmerged.rename(columns=reversed_mapper)
            dfmerged = pd.concat([self.dfmaster, dfmerged], ignore_index=True)
            self.dfmaster = dfmerged

        except (ValueError, FileNotFoundError) as e:
            error_msg = f"{e}\n Files are already merged or not found in the folder {folder_path}"
            self.logger.error(error_msg)
            raise CustomException(error_msg)

    def write_to_excel(self, df, sheet_name=0, overwrite=False):
        try:
            excel_file = self.path
            with xw.App(visible=False) as app:
                with xw.Book(excel_file) as book:
                    sheet = book.sheets[sheet_name]

                    if overwrite == True:
                        last_row = 0
                        sheet.range(f"B{last_row+1}:Z1000").clear_contents()
                        sheet.range(f"B{last_row+1}").options(index=True, header=True).value = df
                    else:
                        last_row = sheet.api.Cells(sheet.api.Rows.Count, "B").End(-4162).Row
                        sheet.range(f"B{last_row+1}").options(index=True, header=False).value = df
                    book.save()

        except Exception as e:
            error_msg = f"Error in writing to excel file {excel_file} : {e}"
            raise CustomException(error_msg)

    def update_new_list(self, folder_path):
        self.merge_excel(folder_path)
        self.write_to_excel(self.dfmaster, overwrite=True)
