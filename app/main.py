from PyQt6 import QtWidgets
from PyQt6.QtCore import QThreadPool
from PyQt6.QtGui import QPixmap

from modules.data_filler import DataFiller
from modules.excel_data_manager import ExcelDataManager
from gui.main_window import Ui_MainWindow
from modules.thread_worker import ThreadWorker

app_version = "v1.1 - alpha"

class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, *args, obj=None, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        self.setupUi(self)

        self.setWindowTitle(f"MasterPlanner {app_version}")

        self.excel_icon_pixmap = QPixmap('gui/resources/icons/excel_icon.png')
        self.datafiller_dst_excel_icon.setPixmap(self.excel_icon_pixmap.scaled(20, 20))
        self.datafiller_src_excel_icon.setPixmap(self.excel_icon_pixmap.scaled(20, 20))

        self.threadpool = QThreadPool()

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
                self.src_file_path_ledit.setText((response[0]))

            if (self.sender() is self.dst_browse_file_button):
                self.dst_file_path_ledit.setText((response[0]))

    def _enable_all_pushbuttons(self, enable):
        pushbutton_object_list = self.tabWidget.findChildren(QtWidgets.QPushButton)
        for push_button in pushbutton_object_list:
            push_button.setEnabled(enable)

    def _check_missing_inputs(self):
        if (self.sender() is self.process_data_filler_button):
                missing_input = False

                line_edit_object_list = self.data_filler_tab.findChildren(QtWidgets.QLineEdit)
                for line_edit in line_edit_object_list:
                    if (line_edit.text() == ""):
                        missing_input = True
                        break
                
                return missing_input

    def _process_data_filler(self, progress_callback=None):
        print("first checkpoint")
        src_excel = ExcelDataManager(self.src_file_path_ledit.text(), 
                                        self.src_sheet_name_ledit.text(), 
                                        int(self.src_col_title_row_ledit.text())-1)
        
        dst_excel = ExcelDataManager(self.dst_file_path_ledit.text(),
                                                self.dst_sheet_name_ledit.text(),
                                                int(self.dst_col_title_row_ledit.text())-1)
        
        src_excel.read_excel()
        dst_excel.read_excel()

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

                worker = ThreadWorker(self._process_data_filler)
                worker.signals.progress.connect(self._update_progress_bar)
                worker.signals.finished.connect(self._thread_complete)

            self.progress_bar = QtWidgets.QProgressDialog("Processing...", "Abort", 0, 100, self)
            self.progress_bar.setMinimumDuration(0)
            self.threadpool.start(worker)

    def _update_progress_bar(self, n):
        self.progress_bar.setValue(n)

    def _thread_complete(self):
        self.progress_bar.close()
        self.progress_bar.destroy()
        self._enable_all_pushbuttons(True)

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = MainWindow()
    window.show()
    app.exec()

