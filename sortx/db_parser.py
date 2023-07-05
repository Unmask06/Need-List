import os
import sqlite3

import pandas as pd

from .master_index import CustomException, MasterIndex


class MiDbParser(MasterIndex):
    def __init__(self, config_file_path="config.xlsm"):
        super().__init__()
        self.config_db(config_file_path)
        self.load_db()

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

    def load_db(self):
        if "database.db" in os.listdir():
            self.logger.info("Database found")
            conn = sqlite3.connect("database.db")
            self.db = pd.read_sql_query("SELECT * FROM database", conn)
        else:
            self.logger.info("Database not found, creating new database")
            self.create_db()
            conn = sqlite3.connect("database.db")
            self.db = pd.read_sql_query("SELECT * FROM database", conn)
            

    def create_db(self):
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
