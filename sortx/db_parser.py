import os
import sqlite3

import pandas as pd

from .master_index import CustomException, MasterIndex


class MiDbParser(MasterIndex):
    def __init__(self, config_file_path):
        super().__init__()
        self.config_db(self.config_file_path)
        self._load_db()

    def config_db(self, config_file_path):
        dfdb = pd.read_excel(config_file_path, sheet_name="database", header=0)
        self.dbconfig = {}
        for i in range(len(dfdb)):
            key = dfdb.iloc[i, 0]
            value = dfdb.iloc[i, 1]
            if isinstance(key, str):
                self.dbconfig[key] = value

        self.dbmapper = {}
        for i in range(len(dfdb)):
            key = dfdb.iloc[i, 3]
            value = dfdb.iloc[i, 4]
            if isinstance(key, str):
                self.dbmapper[key] = value

    def _load_db(self):
        try:
            if "database.db" in os.listdir():
                self.logger.info("Database found")
                conn = sqlite3.connect("database.db")
                self.db = pd.read_sql_query("SELECT * FROM database", conn)
                conn.close()
            else:
                self.logger.info("Database not found, creating new database")
                self._create_db()
                conn = sqlite3.connect("database.db")
                self.db = pd.read_sql_query("SELECT * FROM database", conn)
                conn.close()
        except Exception as e:
            error_msg = f"Error occurred while loading database: {e}"
            self.logger.error(error_msg)
            raise CustomException(error_msg)

    def _create_db(self):
        try:
            self.logger.info("Creating database")
            sheet_name = (
                self.dbconfig["sheet_name"] - 1
                if isinstance(self.dbconfig["sheet_name"], int)
                else self.dbconfig["sheet_name"]
            )
            header = self.dbconfig["header_row_number"] - 1
            df = pd.read_excel(self.dbconfig["database_path"], sheet_name=sheet_name, header=header)
            conn = sqlite3.connect("database.db")
            df.to_sql(name="database", con=conn, if_exists="replace", index=False)
            conn.close()

        except Exception as e:
            error_msg = f"Error occurred while creating database: {e}"
            self.logger.error(error_msg)
            raise CustomException(error_msg)

    def fill_missing_data(self):
        try:
            merge_cols = [col for col in self.dbmapper if col != "doc_no"]

            joined_df = self.preprocess_the_db()

            for col in merge_cols:
                mask = (self.dfmaster[col].isna()) | (self.dfmaster[col] == "")
                self.dfmaster.loc[mask, col] = joined_df.loc[mask, self.dbmapper[col]]

            self.dfmaster = self.dfmaster[self.required_columns]
            self.logger.info("Missing data filled from database")

        except KeyError as ke:
            error_msg = f"Check the database mapping in config file: {ke}"
            self.logger.error(error_msg)
            raise CustomException(error_msg)

        except Exception as e:
            self.logger.error("Error in filling missing data: {}".format(str(e)))

    def preprocess_the_db(self):
            mod_dfmaster = self.dfmaster.copy()
            mod_dfmaster["doc_no"] = mod_dfmaster["doc_no"].apply(lambda x: str(x).replace("/", ""))
            mod_db = self.db.copy()
            mod_db[self.dbmapper["doc_no"]] = mod_db[self.dbmapper["doc_no"]].apply(lambda x: str(x).replace("/", ""))

            joined_df = mod_dfmaster.merge(mod_db, how="left", left_on="doc_no", right_on=self.dbmapper["doc_no"])
            return joined_df