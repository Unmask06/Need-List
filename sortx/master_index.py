import logging
import os
import shutil
import traceback
from datetime import datetime

import pandas as pd
import xlwings as xw


class CustomException(BaseException):
    pass


class MasterIndex:
    def __init__(self, config_file_path="config.xlsm", overwrite_log=True):
        self.log_file = "sortx.log"
        self.config_file_path = config_file_path
        self.setup_logging(overwrite_log)
        self.load_config()
        self.load_mapper()
        self.load_required_columns()
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

    def load_config(self):
        dfconfig = pd.read_excel(self.config_file_path, sheet_name="config", header=0).fillna("")
        self.config = dict(zip(dfconfig.iloc[:, 0], dfconfig.iloc[:, 1]))

    def load_mapper(self):
        dfmapper = pd.read_excel(self.config_file_path, sheet_name="mapper", header=0)
        self.mapper = dict(zip(dfmapper.iloc[:, 0], dfmapper.iloc[:, 1]))
        self.mandate_columns = list(self.mapper.values())

    def load_required_columns(self):
        dfrequired = pd.read_excel(self.config_file_path, sheet_name="field", header=0)
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
        
    def open_master_index(self):
        xw.Book(self.path)

