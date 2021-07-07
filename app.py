from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QDialog
from PyQt5.uic import loadUi
from time import time
import sys

from processors import ProcessDropsEmail, ProcessDrop

# Set application ID
import ctypes
myappid = 'com.ciservicesnow.sscobaxterdrops.1.0'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


class AboutDialog(QDialog):
    def __init__(self) -> None:
        super(AboutDialog, self).__init__()
        loadUi("ui\\about.ui", self)
        self.setFixedSize(self.size())
        self.lblLogo.setScaledContents(True)
        self.lblLogo.setPixmap(QtGui.QPixmap("./assets/ssco_logo.png"))

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super(MainWindow, self).__init__()
        loadUi("ui\mainwindow.ui", self)
        #self.tableWidget.dragEnterEvent = self.dragEnterEvent
        #self.tableWidget.dragMoveEvent = self.dragMoveEvent
        self.tableWidget.dropEvent = self.dropEvent

        self.actionExit.triggered.connect(self.exitApp)
        self.actionAbout.triggered.connect(self.showAbout)

    def exitApp(self):
        sys.exit(app.exec_())

    def showAbout(self):
        dialog = AboutDialog()
        dialog.exec_()

    def loadData(self, file):

        self.drop_data = ProcessDropsEmail(input_file=file)
        self.updateSummarylabels()

        # Populate the tableWidget with the data
        self.tableWidget.setRowCount(len(self.drop_data["successess"]))
        for row, record in enumerate(self.drop_data["successess"].values()):
            for col, item in enumerate(record.values()):
                if col < self.tableWidget.columnCount():
                    self.tableWidget.setItem(row, col, QtWidgets.QTableWidgetItem(item))

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()


        if self.drop_data["failures"]:

            # Populate the failures list with the booking IDs    
            for item in self.drop_data["failures"].values():
                self.lstFailures.addItem(QtWidgets.QListWidgetItem(item["ReferenceNumber"]))

            # Set the summary label red and bold
            self.lblFailures.setStyleSheet("font-weight: bold; color: rgb(255, 0, 0);")

    def updateSummarylabels(self):
        # Set the summary text
        self.lblSender.setText(self.drop_data["sender"])
        self.lblTo.setText(', '.join(self.drop_data["to"]) if len(self.drop_data["to"]) > 1 else self.drop_data["to"][0])
        self.lblSubject.setText(self.drop_data["subject"])
        self.lblNumDrops.setText(str(self.drop_data["numDrops"]))
        self.lblFailures.setText(str(len(self.drop_data["failures"])))
        self.lblAmiaDrops.setText(str(self.drop_data["amiaDrops"]))
        self.lblBaxDrops.setText(str(self.drop_data["baxDrops"]))

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
            self.loadData(event.mimeData().urls()[0].toLocalFile())
            print(f'{time()-bf}')

        else:
            event.ignore()

    def setRecordText(self, record):
        self.txtShipper.setText(record["shipperRaw"])
        self.txtConsignee.setText(record["consigneeRaw"])
        self.txtPickupLocation.setText(record["pickupLocationRaw"])

    def loadRecord(self, row: int, col: int) -> None:

        reference_number = self.tableWidget.item(row,0).text()
        record = self.drop_data["successess"][reference_number]
        self.setRecordText(record)

    def loadFailure(self, item: QtWidgets.QListWidgetItem) -> None:

        reference_number = item.text()
        record = self.drop_data["failures"][reference_number]
        self.setRecordText(record)
        self.txtConsignee.setEnabled(True)
        self.btnSubmitError.setEnabled(True)


    def addRow(self):
        ## Parse text from shipper, consignee, and pickuploactions

        reference_number = self.lstFailures.selectedItems()[0].text()
        record = self.drop_data["failures"][reference_number]

        try:
            record["shipperRaw"] = self.txtShipper.toPlainText()
            record["consigneeRaw"] = self.txtConsignee.toPlainText()
            record["pickupLocationRaw"] = self.txtPickupLocation.toPlainText()
            record = ProcessDrop(record)
            self.drop_data["successess"][record["ReferenceNumber"]] = record
            if record.get("SKU") == "DROP - AMIA":
                self.drop_data["amiaDrops"] += 1
            else:
                self.drop_data["baxDrops"] += 1

        except Exception as e:
            print(e)
            return

        new_row = self.tableWidget.rowCount()
        self.tableWidget.insertRow(new_row)
        for col, item in enumerate(record.values()):
            if col < self.tableWidget.columnCount():
                self.tableWidget.setItem(new_row, col, QtWidgets.QTableWidgetItem(item))

        # Remove failure record from failure list
        self.lstFailures.takeItem(self.lstFailures.selectedIndexes()[0].row())

        # remove failure from failures
        del self.drop_data["failures"][reference_number]

        # test failure label text with number of failures left
        self.lblFailures.setText(str(len(self.drop_data["failures"])))
        
        # set the failure label text back to black and disabled submit button if no failures left
        if not self.drop_data["failures"]:
            self.lblFailures.setStyleSheet("font-weight: normal; color: rgb(0, 0, 0);")
            self.btnSubmitError.setEnabled(False)


        self.updateSummarylabels()

    def exportData(self):

        records = []
        for row in range(self.tableWidget.rowCount()):
            record = [self.tableWidget.item(row, col).text() for col in range(self.tableWidget.columnCount())]
            records.append(record)
        
        print(records)
        



if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon('assets/ssco_icon.ico'))
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
