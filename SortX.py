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
            root_logger = logging.getLogger()
            for handler in root_logger.handlers:
                root_logger.removeHandler(handler)

        file_handler = logging.FileHandler(self.log_file)
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        logging.getLogger().addHandler(file_handler)

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
                logging.info(f"Master index file found at {self.path}")
            elif self.path == "":
                error_msg = "Master index path not specified in config file"
                logging.error(error_msg)
                raise CustomException(error_msg)

        except (FileNotFoundError, TypeError) as e:
            error_msg = f"Master index file not found at {self.path}"
            logging.error(error_msg)
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
                logging.error(error_msg)
                raise ValueError(error_msg)
            else:
                self.dfmaster = dfmaster

        except Exception as e:
            logging.error(e)
            raise e

    def merge_excel(self, folder_path):
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
            dfmerged = pd.concat([self.dfmaster, dfmerged], ignore_index=True)
            self.dfmaster = dfmerged

        except (ValueError, FileNotFoundError) as e:
            logging.error(e)
            print(e)

    def write_to_excel(self, df, sheet_name=0, overwrite=False):
        try:
            app = xw.App(visible=False)
            excel_file = self.path
            book = xw.Book(excel_file)
            sheet = book.sheets[sheet_name]

            if overwrite == True:
                last_row = 0
                sheet.range(f"B{last_row+1}:Z1000").clear_contents()
                sheet.range(f"B{last_row+1}").options(index=True, header=True).value = df
            else:
                last_row = sheet.api.Cells(sheet.api.Rows.Count, "B").End(-4162).Row
                sheet.range(f"B{last_row+1}").options(index=True, header=False).value = df
            book.save()
            book.close()
            app.quit()

        except Exception as e:
            print(traceback.format_exc())

    def update_master_index(self):
        self.write_to_excel(self.dfmaster, overwrite=True)

