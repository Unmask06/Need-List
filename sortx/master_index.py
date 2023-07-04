import logging
import os
import shutil
import traceback
from datetime import datetime

import pandas as pd
import xlwings as xw

from .lister import MasterIndex_Lister
from .db_parser import DB_Parser


class CustomException(BaseException):
    pass

class MasterIndex(MasterIndex_Lister, DB_Parser):

    def __init__(self, config_file_path="config.xlsm", overwrite_log=True):
        super().__init__(config_file_path, overwrite_log)