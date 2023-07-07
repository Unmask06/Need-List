from PyQt6 import QtCore, uic
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QMainWindow,
    QProgressBar,
    QPushButton,
    QTextBrowser,
)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi("gui/ui/main_window.ui", self)
        self.set_defaults()
        self.connect_buttons()

    def set_defaults(self):
        default_values = {
            "tbrowse_Config": "config.xlsm",
            "tbrowse_Merge": r"Need Lists\NL 1\ELECTRICAL",
            "tbrowse_FilePath": r"C:\Users\IDM252577\Desktop\Python Projects\Utility\Need List\Received Files\b1",
        }
        for widget_name, value in default_values.items():
            widget = self.findChild(QTextBrowser, widget_name)
            widget.setText(value)

    def connect_buttons(self):
        buttons = {
            "pb_Config": {"dialog": QFileDialog.getOpenFileName, "widget": "tbrowse_Config"},
            "pb_Merge": {"dialog": QFileDialog.getExistingDirectory, "widget": "tbrowse_Merge"},
            "pb_FilePath": {
                "dialog": QFileDialog.getExistingDirectory,
                "widget": "tbrowse_FilePath",
            },
        }
        for button_name, button_info in buttons.items():
            button = self.findChild(QPushButton, button_name)
            widget = self.findChild(QTextBrowser, button_info["widget"])
            dialog_func = button_info["dialog"]

            button.clicked.connect(
                lambda btn=button, dlg=dialog_func, wgt=widget: self.open_dialog(dlg, wgt)
            )

    def open_dialog(self, dialog_func, widget):
        if dialog_func == QFileDialog.getExistingDirectory:
            folder_path = dialog_func(self, caption="Select Folder", directory="")
            if folder_path:
                widget.setText(folder_path)
        elif dialog_func == QFileDialog.getOpenFileName:
            file_path = dialog_func(
                self,
                caption="Select File",
                directory="",
                filter="Excel Files (*.xlsm *.xlsx *.xls)",
            )
            if file_path:
                widget.setText(file_path[0])

    def show_progress_bar(self):
        self.progressBar = QProgressBar()
        self.progressBar.setRange(0, 100)
        self.progressBar.show()
        self.progressBar.setValue(0)

    def increase_progress(self, value, max_value):
        current_value = int(value / max_value * 100)
        self.progressBar.setValue(current_value)
        if current_value == 100:
            self.progressBar.hide()