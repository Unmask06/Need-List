import os
from datetime import datetime

import pandas as pd
import xlwings as xw

from .master_index import CustomException, MasterIndex


class MiLister(MasterIndex):
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

    def update_new_list(self, folder_path):
        self.merge_excel(folder_path)
        self.write_to_excel(self.dfmaster, overwrite=True)

    def update_folder_link(self, folder_path):
        try:
            for root, dirs, files in os.walk(folder_path):
                for doc_no in dirs:
                    if doc_no in self.dfmaster["doc_no"].values.any():
                        self.dfmaster.loc[
                            self.dfmaster["doc_no"] == doc_no, "source_path"
                        ] = os.path.join(root, doc_no)
                        self.dfmaster.loc[
                            self.dfmaster["doc_no"] == doc_no, "received_status"
                        ] = "closed"
                        self.dfmaster.loc[
                            self.dfmaster["doc_no"] == doc_no, "processed_date"
                        ] = datetime.now().date()
                    else:
                        entry_path = os.path.join(root, doc_no)
                        for sub_root, directories, files in os.walk(entry_path):
                            if not directories:
                                self.dfmaster.loc[len(self.dfmaster)] = {
                                    "doc_no": sub_root.split("\\")[-1],
                                    "source_path": sub_root,
                                    "received_status": "closed",
                                    "imported_from": "extra files",
                                    "processed_date": datetime.now().date(),
                                }
        except Exception as e:
            error_msg = f"{e}\nError while updating folder link"
            self.logger.error(error_msg)
            raise CustomException(error_msg)
