import sys
from excel_create import ExcelCreate

from PySide2.QtWidgets import *

from ui import CreateExcelView


class CreateExcelController(CreateExcelView):
    def __init__(self):
        super().__init__()
        self.model = ExcelCreate()

        # button clicked event
        self.btn_browse.clicked.connect(self.btn_browse_clicked)
        self.btn_create.clicked.connect(self.btn_create_clicked)
        self.btn_cancel.clicked.connect(self.btn_cancel_clicked)

    def btn_browse_clicked(self):
        dialog = QFileDialog()
        dialog.setDirectory('/TD/show')
        dir_path = dialog.getExistingDirectory()
        self.line_path.setText(dir_path)
        self.model.input_path = dir_path

    def btn_create_clicked(self):
        # self.model
        self.message_box()

    def btn_cancel_clicked(self):
        self.line_path.clear()


if __name__ == "__main__":
    app = QApplication()
    cc = CreateExcelController()
    sys.exit(app.exec_())
