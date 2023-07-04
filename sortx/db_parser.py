


from .setup import MasterIndex_Setup , CustomException
from .lister import MasterIndex_Lister

class DB_Parser(MasterIndex_Setup):

    def __init__(self, db_file_path="database.db"):
        super().__init__(db_file_path)
        self.read_database(db_file_path)

    def read_database(self, db_file_path):
        print("read_database")

