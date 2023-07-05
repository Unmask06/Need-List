import sqlite3
import pandas as pd


from .master_index import MasterIndex , CustomException

class MiDbParser(MasterIndex):

    def __init__(self, db_file_path="database.db"):
        super().__init__(db_file_path)
        self.read_database(db_file_path)

    def read_database(self, db_file_path):
        print("read_database")

