import sys
import traceback
import time

from PyQt6 import QtWidgets

from gui import MainWindow
from sortx import CustomException, MasterIndex, MiDbParser, MiLister

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = MainWindow()

    widget = {
        "config": window.tbrowse_Config,
        "xl_folder_path": window.tbrowse_Merge,
        "file_path": window.tbrowse_FilePath,
    }
    paths = {"config": "", "xl_folder_path": "", "file_path": ""}

    def get_path(field, text_box_dict=widget):
        for key, widget in text_box_dict.items():
            if field == key:
                path = widget.toPlainText()
                return path

    def set_path(field):
        paths[field] = get_path(field)

    def run_merge():
        try:
            #! simplify this in next rev , set path in one go
            for key, value in paths.items():
                set_path(key)

            lister = MiLister(config_file_path=paths["config"])
            lister.merge_excel(paths["xl_folder_path"])
            lister.write_to_excel()

            lister.logger.info("Done! Files merged.")

        except CustomException as e:
            print(e)
            # print(traceback.format_exc())

    def run_update_folder_link():
        try:
            for key, value in paths.items():
                set_path(key)

            lister = MiLister(config_file_path=paths["config"])
            lister.update_folder_link(paths["file_path"])
            lister.write_to_excel()

            lister.logger.info("Done! File Path updated.")

        except CustomException as e:
            print(e)
            print(traceback.format_exc())

    def run_open_master_index():
        for key, value in paths.items():
            set_path(key)
        lister = MiLister(config_file_path=paths["config"])
        lister.open_master_index()

    def run_fill_data():
        for key, value in paths.items():
            set_path(key)
        db_parser = MiDbParser(config_file_path=paths["config"])
        db_parser.fill_missing_data()
        db_parser.write_to_excel()

    window.show()
    window.pb_RunMerge.clicked.connect(run_merge)
    window.pb_RunFilePath.clicked.connect(run_update_folder_link)
    window.pb_RunOpenMi.clicked.connect(run_open_master_index)
    window.pb_RunFillData.clicked.connect(run_fill_data)
    sys.exit(app.exec())

#! Add progress bar for filling data