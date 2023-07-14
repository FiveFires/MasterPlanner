"""
MasterPlanner App

Description: This module is the application entry point. 
             It defines the main window and its functionality for a PyQt6-based application, 
             including tabs for data filling and welding planning, file browsing,
             long-running process execution, and progress updates.

Author: Adam Ondryas
Email: adam.ondryas@gmail.com

This software is distributed under the GPL v3.0 license.
"""

# external module imports
from PyQt6 import QtWidgets
from PyQt6.QtCore import QThreadPool, QSize
from PyQt6.QtGui import QPixmap, QIcon

# local module imports
from modules.data_filler import DataFiller
from modules.welding_planner import WeldingPlanner
from modules.excel_data_manager import ExcelDataManager
from gui.main_window import Ui_MainWindow
from modules.thread_worker import ThreadWorker

app_version = "v1.3.1"

class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, *args, obj=None, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        self.setupUi(self)

        self.setFixedSize(QSize(420, 540))
        self.setWindowTitle(f"MasterPlanner {app_version}")
        self.excel_icon_pixmap = QPixmap('gui/resources/icons/excel_icon.png')

        self._setup_welding_planner_tab()
        self._setup_data_filler_tab()

        self.threadpool = QThreadPool()

    def _setup_welding_planner_tab(self):
        self.welding_planner_excel_icon.setPixmap(self.excel_icon_pixmap.scaled(20, 20))
        self.bi_reser_excel_icon.setPixmap(self.excel_icon_pixmap.scaled(20, 20))
        self.manuf_plan_excel_icon.setPixmap(self.excel_icon_pixmap.scaled(20, 20))
        self.batch_data_excel_icon.setPixmap(self.excel_icon_pixmap.scaled(20, 20))
        self.welding_planner_col_title_row_ledit.setText("1")
        self.bi_reser_col_title_row_ledit.setText("1")
        self.manuf_plan_col_title_row_ledit.setText("1")
        self.batch_data_col_title_row_ledit.setText("2")

        self.welding_planner_browse_file_button.clicked.connect(self._get_file_path)
        self.welding_planner_fpath_ledit.textChanged.connect(self._update_button_text)
        
        self.bi_reser_browse_file_button.clicked.connect(self._get_file_path)
        self.manuf_plan_browse_file_button.clicked.connect(self._get_file_path)
        self.batch_data_browse_file_button.clicked.connect(self._get_file_path)
        self.process_welding_planner_button.clicked.connect(self._long_process_threadcall)

    def _setup_data_filler_tab(self):
        self.datafiller_dst_excel_icon.setPixmap(self.excel_icon_pixmap.scaled(20, 20))
        self.datafiller_src_excel_icon.setPixmap(self.excel_icon_pixmap.scaled(20, 20))

        self.src_browse_file_button.clicked.connect(self._get_file_path)
        self.dst_browse_file_button.clicked.connect(self._get_file_path)
        self.process_data_filler_button.clicked.connect(self._long_process_threadcall)

    def _get_file_path(self):
        file_filter = 'Excel File (*.xlsx *.xls)'
        
        response = QtWidgets.QFileDialog.getOpenFileName(
            parent=self,
            caption='Select a file',
            filter=file_filter,
            initialFilter=file_filter
        )

        if response != "":
            if (self.sender() is self.src_browse_file_button):
                self.src_fpath_ledit.setText((response[0]))

            elif (self.sender() is self.dst_browse_file_button):
                self.dst_fpath_ledit.setText((response[0]))

            elif (self.sender() is self.welding_planner_browse_file_button):
                self.welding_planner_fpath_ledit.setText((response[0]))

            elif (self.sender() is self.bi_reser_browse_file_button):
                self.bi_reser_fpath_ledit.setText((response[0]))

            elif (self.sender() is self.manuf_plan_browse_file_button):
                self.manuf_plan_fpath_ledit.setText((response[0]))

            elif (self.sender() is self.batch_data_browse_file_button):
                self.batch_data_fpath_ledit.setText((response[0]))

    def _enable_all_pushbuttons(self, enable):
        pushbutton_object_list = self.tabWidget.findChildren(QtWidgets.QPushButton)
        for push_button in pushbutton_object_list:
            push_button.setEnabled(enable)

    def _check_missing_inputs(self):
        missing_input = False

        if (self.sender() is self.process_data_filler_button):
            line_edit_object_list = self.data_filler_tab.findChildren(QtWidgets.QLineEdit)

        elif(self.sender() is self.process_welding_planner_button):
            line_edit_object_list = self.welding_planner_tab.findChildren(QtWidgets.QLineEdit)
            line_edit_object_list.remove(self.welding_planner_fpath_ledit)

        for line_edit in line_edit_object_list:
            if (line_edit.text() == ""):
                missing_input = True
                break

        return missing_input

    def _update_button_text(self):
        if(self.welding_planner_fpath_ledit.text() != ""):
            self.process_welding_planner_button.setText("Update Welding Plan")
        else:
            self.process_welding_planner_button.setText("Create Welding Plan")

    def _run_welding_planner(self, progress_callback=None):
        
        bi_reservations_excel = ExcelDataManager(self.bi_reser_fpath_ledit.text(),
                                                 sheet_name=0,
                                                 column_name_row=int(
                                                 self.bi_reser_col_title_row_ledit.text())-1)

        manufacturing_plan_excel = ExcelDataManager(self.manuf_plan_fpath_ledit.text(),
                                                    sheet_name=0,
                                                    column_name_row=int(
                                                    self.manuf_plan_col_title_row_ledit.text())-1)

        batch_database_excel = ExcelDataManager(self.batch_data_fpath_ledit.text(),
                                                sheet_name=0,
                                                column_name_row=int(
                                                self.batch_data_col_title_row_ledit.text())-1)

        if(self.welding_planner_fpath_ledit.text() != ""):
            welding_planner_excel = ExcelDataManager(self.welding_planner_fpath_ledit.text(),
                                                     sheet_name="Welding Plan",
                                                     column_name_row=int(
                                                     self.welding_planner_col_title_row_ledit.text())-1) 
        else:
            welding_planner_excel = None

        welding_planner_instance = WeldingPlanner(welding_planner_excel)
        
        welding_planner_instance.plan_welding(bi_reservations_excel, manufacturing_plan_excel, 
                                              batch_database_excel, progress_callback=progress_callback)

    def _run_data_filler(self, progress_callback):
        src_excel = ExcelDataManager(self.src_fpath_ledit.text(), 
                                        self.src_sheet_name_ledit.text(), 
                                        int(self.src_col_title_row_ledit.text())-1)
        
        dst_excel = ExcelDataManager(self.dst_fpath_ledit.text(),
                                                self.dst_sheet_name_ledit.text(),
                                                int(self.dst_col_title_row_ledit.text())-1)
        
        src_excel.df = src_excel.read_excel()
        dst_excel.df = dst_excel.read_excel()

        data_filler_instance = DataFiller(src_excel, dst_excel, 
                                            self.src_lookup_column_ledit.text(), 
                                            self.src_copy_column_ledit.text(), 
                                            self.dst_lookup_column_ledit.text(), 
                                            self.dst_fill_column_ledit.text())

        data_filler_instance.fill_data(progress_callback=progress_callback)

    def _long_process_threadcall(self):
        if self._check_missing_inputs() == False:

            self._enable_all_pushbuttons(False)

            if (self.sender() is self.process_data_filler_button):

                worker = ThreadWorker(self._run_data_filler)

            elif(self.sender() is self.process_welding_planner_button):
                worker = ThreadWorker(self._run_welding_planner)

            worker.signals.progress.connect(self._update_progress_bar)
            worker.signals.error.connect(self._thread_raised_exception)
            worker.signals.result.connect(self._thread_processed_successfully)
            worker.signals.finished.connect(self._thread_finished)

            self.progress_bar = QtWidgets.QProgressDialog("Processing...", "Abort", 0, 100, self)
            self.progress_bar.setMinimumDuration(0)
            self.threadpool.start(worker)

    def _update_progress_bar(self, n):
        self.progress_bar.setValue(n)

    def _thread_processed_successfully(self):
        msgBox = QtWidgets.QMessageBox()
        msgBox.setWindowIcon(QIcon('gui/resources/icons/master_planner_icon.png'))
        msgBox.setText("Processing is complete!")
        msgBox.setWindowTitle("MasterPlanner Processing")
        msgBox.setStyleSheet("QLabel{min-width: 200px; min-height: 100px;}")
        msgBox.exec()

    def _thread_raised_exception(self):
        msgBox = QtWidgets.QMessageBox()
        msgBox.setWindowIcon(QIcon('gui/resources/icons/master_planner_icon.png'))
        msgBox.setText("Processing failed, unknown error!\nCheck the inputs.")
        msgBox.setWindowTitle("MasterPlanner Processing")
        msgBox.setStyleSheet("QLabel{min-width: 200px; min-height: 100px;}")
        msgBox.exec()

        self.progress_bar.close()

    def _thread_finished(self):
        self._enable_all_pushbuttons(True)

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = MainWindow()
    window.show()
    app.exec()

