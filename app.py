import enum
from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.uic import loadUi
from datetime import datetime
from time import time
import sys

from processors import ProcessDrop


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super(MainWindow, self).__init__()
        loadUi("ui\mainwindow.ui", self)
        #self.tableWidget.dragEnterEvent = self.dragEnterEvent
        #self.tableWidget.dragMoveEvent = self.dragMoveEvent
        self.tableWidget.dropEvent = self.dropEvent

    def loadData(self, file):

        self.drop_data = ProcessDrop(input_file=file, debug=False)
        self.lblSender.setText(self.drop_data["sender"])


        self.tableWidget.setRowCount(len(self.drop_data["successess"]))
        for row, record in enumerate(self.drop_data["successess"]):
            self.tableWidget.setItem(row, 0, QtWidgets.QTableWidgetItem(record["ReferenceNumber"]))

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls:
            event.accept()
        else:
            event.ignore()
    
    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls:
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls:
            event.setDropAction(Qt.CopyAction)
            event.accept()

            bf = time()
            self.drop_data = ProcessDrop(event.mimeData().urls()[0].toLocalFile())
            print(f'{time()-bf}')

            self.lblSender.setText(self.drop_data["sender"])
            self.lblTo.setText(', '.join(self.drop_data["to"]) if len(self.drop_data["to"]) > 1 else self.drop_data["to"])
            self.lblSubject.setText(self.drop_data["subject"])
            self.lblNumDrops.setText(str(self.drop_data["numDrops"]))
            self.lblFailures.setText(str(len(self.drop_data["failures"])))
            self.lblAmiaDrops.setText(str(self.drop_data["amiaDrops"]))
            self.lblBaxDrops.setText(str(self.drop_data["baxDrops"]))



            self.tableWidget.setRowCount(len(self.drop_data["successess"]))
            for row, record in enumerate(self.drop_data["successess"].values()):
                for col, item in enumerate(record.values()):
                    if col < self.tableWidget.columnCount():
                        self.tableWidget.setItem(row, col, QtWidgets.QTableWidgetItem(item))

            self.tableWidget.resizeColumnsToContents()
        else:
            event.ignore()

    def setRecordText(self, record):
        self.txtShipper.setText(record["shipperRaw"])
        self.txtConsignee.setText(record["consigneeRaw"])
        self.txtPickupLocation.setText(record["pickupLocationRaw"])

    def loadRecord(self, row, col):

        reference_number = self.tableWidget.item(row,0).text()
        record = self.drop_data["successess"][reference_number]
        self.setRecordText(record)

    def exportData(self):

        records = []
        for row in range(self.tableWidget.rowCount()):
            record = [self.tableWidget.item(row, col).text() for col in range(self.tableWidget.columnCount())]
            records.append(record)
        
        print(records)
        



if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
