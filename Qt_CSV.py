#!/usr/bin/env python
#-*- coding:utf-8 -*-
import csv, codecs 
import sip
import os
sip.setapi('QString', 2)
sip.setapi('QVariant', 2)

from PyQt4 import QtGui, QtCore
try:
    from PyQt4.QtCore import QString
except ImportError:
    # we are using Python3 so QString is not defined
    QString = type("")
from PyQt4.QtGui import QApplication, QImage, QPainter, QWidget, QVBoxLayout,\
    QDesktopWidget, QFileDialog, QSizePolicy, QHBoxLayout, QPrinter

class MyWindow(QtGui.QWidget):
    def __init__(self, fileName, parent=None):
        super(MyWindow, self).__init__(parent)
        self.fileName = fileName
        self.fname = "Liste"
        self.model = QtGui.QStandardItemModel(self)

        self.tableView = QtGui.QTableView(self)
        self.tableView.setStyleSheet(stylesheet(self))
        self.tableView.setModel(self.model)
        self.tableView.horizontalHeader().setStretchLastSection(True)
        self.tableView.setShowGrid(True)
        self.tableView.setGeometry(10, 50, 780, 645)
        self.tableView.setSizePolicy(QtGui.QSizePolicy.Preferred,QtGui.QSizePolicy.Preferred)

        self.pushButtonLoad = QtGui.QPushButton(self)
        self.pushButtonLoad.setText("Load CSV")
        self.pushButtonLoad.clicked.connect(self.on_pushButtonLoad_clicked)
        self.pushButtonLoad.setFixedWidth(80)
        self.pushButtonLoad.setStyleSheet(stylesheet(self))
#        self.pushButtonLoad.move(10, 10)

        self.pushButtonWrite = QtGui.QPushButton(self)
        self.pushButtonWrite.setText("Save CSV")
        self.pushButtonWrite.clicked.connect(self.on_pushButtonWrite_clicked)
        self.pushButtonWrite.setFixedWidth(80)
        self.pushButtonWrite.setStyleSheet(stylesheet(self))
#        self.pushButtonWrite.move(100, 10)

        self.pushButtonPreview = QtGui.QPushButton(self)
        self.pushButtonPreview.setText("Print Preview")
        self.pushButtonPreview.clicked.connect(self.handlePreview)
        self.pushButtonPreview.setFixedWidth(80)
        self.pushButtonPreview.setStyleSheet(stylesheet(self))
#        self.pushButtonPreview.move(200, 10)

        self.pushButtonPrint = QtGui.QPushButton(self)
        self.pushButtonPrint.setText("Print")
        self.pushButtonPrint.clicked.connect(self.handlePrint)
        self.pushButtonPrint.setFixedWidth(80)
        self.pushButtonPrint.setStyleSheet(stylesheet(self))
#        self.pushButtonPrint.move(290, 10)

        self.pushAddRow = QtGui.QPushButton(self)
        self.pushAddRow.setText("add Row")
        self.pushAddRow.clicked.connect(self.addRow)
        self.pushAddRow.setFixedWidth(80)
        self.pushAddRow.setStyleSheet(stylesheet(self))
#        self.pushAddRow.move(400, 10)

        self.pushDeleteRow = QtGui.QPushButton(self)
        self.pushDeleteRow.setText("delete Row")
        self.pushDeleteRow.clicked.connect(self.removeRow)
        self.pushDeleteRow.setFixedWidth(80)
        self.pushDeleteRow.setStyleSheet(stylesheet(self))
#        self.pushDeleteRow.move(490, 10)

        self.pushAddColumn = QtGui.QPushButton(self)
        self.pushAddColumn.setText("add Column")
        self.pushAddColumn.clicked.connect(self.addColumn)
        self.pushAddColumn.setFixedWidth(80)
        self.pushAddColumn.setStyleSheet(stylesheet(self))
#        self.pushAddColumn.move(600, 10)

        self.pushDeleteColumn = QtGui.QPushButton(self)
        self.pushDeleteColumn.setText("delete Column")
        self.pushDeleteColumn.clicked.connect(self.removeColumn)
        self.pushDeleteColumn.setFixedWidth(80)
        self.pushDeleteColumn.setStyleSheet(stylesheet(self))
#        self.pushDeleteColumn.move(690, 10)

        grid = QtGui.QGridLayout()
        grid.setSpacing(10)
        grid.addWidget(self.pushButtonLoad, 0, 0)
        grid.addWidget(self.pushButtonWrite, 0, 1)
        grid.addWidget(self.pushAddRow, 0, 2)
        grid.addWidget(self.pushDeleteRow, 0, 3)
        grid.addWidget(self.pushAddColumn, 0, 4)
        grid.addWidget(self.pushDeleteColumn, 0, 5)
        grid.addWidget(self.pushButtonPreview, 0, 6)
        grid.addWidget(self.pushButtonPrint, 0, 7, 1, 1, QtCore.Qt.AlignRight)
        grid.addWidget(self.tableView, 1, 0, 1, 8)
        self.setLayout(grid)

        item = QtGui.QStandardItem()
        self.model.appendRow(item)
        self.model.setData(self.model.index(0, 0), "new", 0)
        self.tableView.resizeColumnsToContents()

    def loadCsv(self, fileName):
        self.model.clear()
        with open(fileName, "rb") as fileInput:
            self.fname = os.path.splitext(str(fileName))[0].split("/")[-1]
            self.setWindowTitle(self.fname)
            reader = csv.reader(fileInput, delimiter = '\t')
            for row in reader:    
                items = [QtGui.QStandardItem(field.decode('utf-8')) for field in row]
                self.model.appendRow(items)

    def writeCsv(self, fileName):
        # find empty cells
        for row in range(self.model.rowCount()):
            for column in range(self.model.columnCount()):
                myitem = self.model.item(row,column)
                if myitem is None:
                    item = QtGui.QStandardItem("")
                    self.model.setItem(row, column, item)
        with open(fileName, "wb") as fileOutput:
            writer = csv.writer(fileOutput, delimiter = '\t')
            for rowNumber in range(self.model.rowCount()):
                fields = [self.model.data(self.model.index(rowNumber, columnNumber),
                        QtCore.Qt.DisplayRole).encode("utf-8")
                    for columnNumber in range(self.model.columnCount())]
                writer.writerow(fields)

    @QtCore.pyqtSlot()
    def on_pushButtonWrite_clicked(self):
		result = QFileDialog.getSaveFileName(self, "CSV speichern",
                                                '/home',
                                                "Tabelle (*.csv *.txt)")
                if result:
                    self.writeCsv(result)

    @QtCore.pyqtSlot()
    def on_pushButtonLoad_clicked(self):
		result = QFileDialog.getOpenFileName(self, "CSV laden",
                                                '/home',
                                                "Tabelle (*.csv *.txt)")#.decode('utf-8')
                if result:
                    self.loadCsv(result)

    def handlePrint(self):
        dialog = QtGui.QPrintDialog()
        if dialog.exec_() == QtGui.QDialog.Accepted:
            self.handlePaintRequest(dialog.printer())

    def handlePreview(self):
        dialog = QtGui.QPrintPreviewDialog()
        dialog.setFixedSize(1000,700)
        dialog.paintRequested.connect(self.handlePaintRequest)
        dialog.exec_()

    def handlePaintRequest(self, printer):
#        printer = QtGui.QPrinter(QtGui.QPrinter.PrinterResolution)
        # find empty cells
        for row in range(self.model.rowCount()):
            for column in range(self.model.columnCount()):
                myitem = self.model.item(row,column)
                if myitem is None:
                    item = QtGui.QStandardItem("")
                    self.model.setItem(row, column, item)
        printer.setDocName(self.fname)
        document = QtGui.QTextDocument()
        cursor = QtGui.QTextCursor(document)
        model = self.tableView.model()
        table = cursor.insertTable(model.rowCount(), model.columnCount())
        for row in range(table.rows()):
            for column in range(table.columns()):
                cursor.insertText(model.item(row, column).text())
                cursor.movePosition(QtGui.QTextCursor.NextCell)
        document.print_(printer)


    def removeRow(self):
        model = self.model
        indices = self.tableView.selectionModel().selectedRows() 
        for index in sorted(indices):
            model.removeRow(index.row()) 

    def addRow(self):
        item = QtGui.QStandardItem("")
        self.model.appendRow(item)


    def removeColumn(self):
        model = self.model
        indices = self.tableView.selectionModel().selectedColumns() 
        for index in sorted(indices):
            model.removeColumn(index.column()) 

    def addColumn(self):
        count = self.model.columnCount()
        print (count)
        self.model.setColumnCount(count + 1)
        self.model.setData(self.model.index(0, count), "new column", 0)
        self.tableView.resizeColumnsToContents()

def stylesheet(self):
        return """
        QTableView
        {
			border: 1px solid grey;
			border-radius: 0px;
			font-size: 12px;
        	background-color: #f8f8f8;
			selection-color: white;
			selection-background-color: #00ED56;
        }

		QTableView QTableCornerButton::section {
    		background: #D6D1D1;
    		border: 1px outset black;
		}

		QPushButton
		{
			font-size: 10px;
			border: 1px inset grey;
			height: 24px;
			width: 80px;
			color: black;
			background-color: #e8e8e8;
			background-position: bottom-left;
		} 

		QPushButton::hover
		{
			border: 1px inset goldenrod;
			font-weight: bold;
			color: #e8e8e8;
			background-color: green;
		} 
	"""

if __name__ == "__main__":
    import sys

    app = QtGui.QApplication(sys.argv)
    app.setApplicationName('MyWindow')
    main = MyWindow('')
#    main.setMaximumSize(800, 700)
    main.setMinimumSize(800, 300)
    main.setGeometry(0,0,800,700)
    main.setWindowTitle("CSV Viewer")
    main.show()

sys.exit(app.exec_())