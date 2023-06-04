"""
Module: ThreadWorker
Description: This module allows long processes to run in different threads 
             so that they do not block the UI thread
             
Author: Adam Ondryas
Email: adam.ondryas@gmail.com
"""

from PyQt6.QtCore import QRunnable, QThreadPool, pyqtSlot, pyqtSignal, QObject

class WorkerSignals(QObject):
    '''
    Defines the signals available from a running worker thread.

    Supported signals are:

    finished
        No data

    progress
        int indicating % progress

    '''
    finished = pyqtSignal()
    progress = pyqtSignal(int)

class ThreadWorker(QRunnable):
    '''
    Worker thread

    Inherits from QRunnable to handler worker thread setup, signals and wrap-up.

    :param callback: The function callback to run on this worker thread. Supplied args and
                     kwargs will be passed through to the runner.
    :type callback: function
    :param args: Arguments to pass to the callback function
    :param kwargs: Keywords to pass to the callback function

    '''

    def __init__(self, fn, *args, **kwargs):
        super(ThreadWorker, self).__init__()

        # Store constructor arguments (re-used for processing)
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

        # Add the callback to kwargs
        self.kwargs['progress_callback'] = self.signals.progress

    @pyqtSlot()
    def run(self):
            self.fn(*self.args, **self.kwargs)

            #emit the finished signal once the thread has finished
            self.signals.finished.emit()