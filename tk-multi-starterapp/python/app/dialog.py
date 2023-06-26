# Copyright (c) 2013 Shotgun Software Inc.
#
# CONFIDENTIAL AND PROPRIETARY
#
# This work is provided "AS IS" and subject to the Shotgun Pipeline Toolkit
# Source Code License included in this distribution package. See LICENSE.
# By accessing, using, copying or modifying this work you indicate your
# agreement to the Shotgun Pipeline Toolkit Source Code License. All rights
# not expressly granted therein are reserved by Shotgun Software Inc.


import sgtk
import os
import sys
import threading
import subprocess

# by importing QT from sgtk rather than directly, we ensure that
# the code will be compatible with both PySide and PyQt.

from sgtk.platform.qt import QtCore, QtGui

from .api import ExcelCreate
# from .api import LoadConvertModel

from .ui.dialog import Ui_Dialog

# standard toolkit logger
logger = sgtk.platform.get_logger(__name__)


def show_dialog(app_instance):
    """
    Shows the main dialog window.
    """
    # in order to handle UIs seamlessly, each toolkit engine has methods for launching
    # different types of windows. By using these methods, your windows will be correctly
    # decorated and handled in a consistent fashion by the system.

    # we pass the dialog class to this method and leave the actual construction
    # to be carried out by toolkit.
    app_instance.engine.show_dialog("IO Manager", app_instance, AppDialog)


class AppDialog(QtGui.QWidget):
    """
    Main application dialog window
    """

    def __init__(self):
        """
        Constructor
        """
        # first, call the base class and let it do its thing.
        QtGui.QWidget.__init__(self)

        # now load in the UI that was created in the UI designer
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.main_stack.setCurrentIndex(0)

        self.api_1 = ExcelCreate()
        # self.api_2 = LoadConvertModel()

        # most of the useful accessors are available through the Application class instance
        # it is often handy to keep a reference to this. You can get it via the following method:
        self._app = sgtk.platform.current_bundle()

        # logging happens via a standard toolkit logger
        logger.info("Launching Starter Application...")

        # via the self._app handle we can for example access:
        # - The engine, via self._app.engine
        # - A Shotgun API instance, via self._app.shotgun
        # - An Sgtk API instance, via self._app.sgtk

        # excel_create
        self.ui.ui1_button_2.clicked.connect(self.create_excel)

        self.ui.btn_browse.clicked.connect(self.btn_browse_clicked)
        self.ui.btn_clear.clicked.connect(self.btn_clear_clicked)
        # self.ui.btn_create.clicked.connect(self.btn_create_clicked)
        # self.ui.btn_create.clicked.connect(self.save_folder_open)
        self.ui.btn_cancel.clicked.connect(self.btn_cancel_clicked)

        # convert
        self.ui.ui2_button_2.clicked.connect(self.convert)

        # self.ui.combo_box.currentIndexChanged.connect(self.combo_box_changed)
        self.ui.browse_button.clicked.connect(self.browse_clicked)
        # self.ui.ok_button.clicked.connect(self.ok_clicked)
        self.ui.cancel_button.clicked.connect(self.btn_cancel_clicked)

    def create_excel(self):
        self.ui.main_stack.setCurrentIndex(2)

    def convert(self):
        self.ui.main_stack.setCurrentIndex(1)

    ## create excel ##
    def btn_browse_clicked(self):
        dialog = QtGui.QFileDialog()
        dialog.setDirectory('/TD/show')
        dir_path = dialog.getExistingDirectory()
        self.ui.line_dir_path.setText(dir_path)

    def btn_clear_clicked(self):
        self.ui.line_dir_path.clear()

    # def btn_create_clicked(self):
    #     self.api_1.excel_create()
    #     self.label_save_path.setText(self.api_1.excel_path)
    #     self.message_box()

    # def save_folder_open(self):
    #     if self.retval == QMessageBox.Open:
    #         subprocess.Popen(['gio', 'open', self.model.excel_path])

    def btn_cancel_clicked(self):
        self.ui.main_stack.setCurrentIndex(0)

    ## convert & upload ##
    # def combo_box_changed(self):
    #     self.selected_project = self.view.combo_box.currentText()
    #     self.view.label.setText(f'Converting Project: {self.selected_project}')
    #     print(f'selected project: {self.selected_project}')

    def browse_clicked(self):
        dialog = QtGui.QFileDialog()
        dialog.setDirectory('/TD/show')
        file_name = dialog.getOpenFileName(self,
                                                 "Select File",
                                                 filter="Excel Files (*.xlsx);;CSV Files (*.csv)")
        self.ui.line_edit.setText(file_name[0])

    # def ok_clicked(self):
    #     if not self.model.file_path:
    #         self.view.message_box('Choose the xlsx file')
    #         self.clean_up()
    #         return
    #     if not self.selected_project:
    #         self.view.message_box('Choose the project')
    #         self.clean_up()
    #         return
    #     if not self.model.file_path and self.selected_project:
    #         self.view.message_box('Choose the xlsx file and project')
    #         self.clean_up()
    #     else:
    #         self.model.set_file_path(self.model.file_path)
    #         self.model.data_info()
    #         self.model.get_project(self.selected_project)
    #         self.model.get_sequence_and_upload()
    #         self.model.get_shot_and_upload()
    #         self.model.video_uploader()
    #         self.view.message_box('Work is Done')
    #
    # def clean_up(self):
    #     self.view.line_edit.setText('')
    #     self.selected_xlsx = ''
    #     self.model.file_path = ''

