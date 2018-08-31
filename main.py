import os
import pandas as pd
import requests
import sys
import xlwings as xw

from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QDate, pyqtSlot
from PyQt5.QtWidgets import (QAction, QApplication, QCalendarWidget, QDateEdit, QGridLayout, QGroupBox, QHBoxLayout, QLabel, QLayout, QMainWindow, QWidget, QPushButton, QLineEdit, QTextEdit, QFileDialog)


class Window(QWidget):
    def __init__(self):
        super(Window, self).__init__()
        self.setWindowTitle("Query Cocorahs")
        self.initUI()


    def initUI(self):
        self.createPreviewGroupBox()
        self.createDatesGroupBox()
        self.infoPanel()

        layout = QGridLayout()
        # layout.addWidget(self.previewGroupBox, 0, 0)
        layout.addWidget(self.datesGroupBox, 0, 1)
        layout.addWidget(self.infoPanelBox, 0, 2, 0, 1)
        layout.setSizeConstraint(QLayout.SetFixedSize)
        self.setLayout(layout)
        self.previewLayout.setRowMinimumHeight(0,
                self.calendar.sizeHint().height())
        self.previewLayout.setColumnMinimumWidth(0,
                self.calendar.sizeHint().width())


    def createPreviewGroupBox(self):
        self.previewGroupBox = QGroupBox("Calendar")
        self.calendar = QCalendarWidget()
        self.calendar.setMaximumDate(QDate(3000, 1, 1))
        self.calendar.setGridVisible(True)
        self.previewLayout = QGridLayout()
        self.previewLayout.addWidget(self.calendar, 0, 0)
        self.previewGroupBox.setLayout(self.previewLayout)


    def createDatesGroupBox(self):
        self.datesGroupBox = QGroupBox(self.tr("Selections"))

        # QUERY
        self.queryButton = QPushButton("Query Database")
        self.queryButton.clicked.connect(self.query_on_click)

        # START DATES
        self.currentDateEdit = QDateEdit()
        self.currentDateEdit.setDisplayFormat('dd MMM yyyy')
        self.currentDateEdit.setDate(self.calendar.selectedDate())
        self.currentDateEdit.setDateRange(self.calendar.minimumDate(),
                self.calendar.maximumDate())
        self.currentDateLabel = QLabel("&Start Date:")
        self.currentDateLabel.setBuddy(self.currentDateEdit)

        # END DATES
        self.maximumDateEdit = QDateEdit()
        self.maximumDateEdit.setDisplayFormat('dd MMM yyyy')
        self.maximumDateEdit.setDate(self.calendar.selectedDate())
        self.maximumDateEdit.setDateRange(self.calendar.minimumDate(),
                self.calendar.maximumDate())
        self.maximumDateLabel = QLabel("&End Date:")
        self.maximumDateLabel.setBuddy(self.maximumDateEdit)

        # LONGITUDE
        self.textboxLong = QLineEdit(self)
        self.textboxLongLabel = QLabel("Longitude:")

        # LATITUDE
        self.textboxLat = QLineEdit(self)
        self.textboxLatLabel = QLabel("Latitude:")

        # STATES
        self.textboxStates = QLineEdit(self)
        self.textboxStatesLabel = QLabel("States: 'CO,NE,WY'")

        dateBoxLayout = QGridLayout()
        dateBoxLayout.addWidget(self.currentDateLabel, 1, 0)
        dateBoxLayout.addWidget(self.currentDateEdit, 1, 1, 1, 2)
        dateBoxLayout.addWidget(self.maximumDateLabel, 2, 0)
        dateBoxLayout.addWidget(self.maximumDateEdit, 2, 1, 1, 2)
        dateBoxLayout.addWidget(self.textboxLongLabel, 3, 0)
        dateBoxLayout.addWidget(self.textboxLong, 3, 1, 1, 2)
        dateBoxLayout.addWidget(self.textboxLatLabel, 4, 0)
        dateBoxLayout.addWidget(self.textboxLat, 4, 1, 1, 2)
        dateBoxLayout.addWidget(self.textboxStatesLabel, 5, 0)
        dateBoxLayout.addWidget(self.textboxStates, 5, 1, 1, 2)
        dateBoxLayout.addWidget(self.queryButton, 6, 0,1,3)
        dateBoxLayout.setRowStretch(5, 1)
        self.datesGroupBox.setLayout(dateBoxLayout)


    def infoPanel(self):
        self.infoPanelBox = QGroupBox(self.tr("Info Console"))
        self.infoConsole = QTextEdit(self)
        infoPanelBoxLayout = QGridLayout()
        infoPanelBoxLayout.addWidget(self.infoConsole, 0, 0, 1, 1)
        self.infoPanelBox.setLayout(infoPanelBoxLayout)



    def query_on_click(self):
        long = self.textboxLong.text()
        lat = self.textboxLat.text()
        startDate = self.currentDateEdit.date()
        endDate = self.maximumDateEdit.date()
        date_range = list(range(startDate.month(), endDate.month() + 1))

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"Choose a database (.csv) to query", "","All Files (*);;Python Files (*.csv)", options=options)
        if fileName:
            df = pd.read_csv(fileName)
            query(long=long, lat=lat, date_range=date_range, states=states)


def query(long, lat, date_range, states):


    # if month_range == []:
    #     window.infoConsole.append("'Start Month' must be before 'End Month'\n")
    # elif job == '' and emp == '':
    #     window.infoConsole.append("No Job or Employee chosen\n")
    # elif job not in listJobs and job != '':
    #     window.infoConsole.append(f"Job number:{job} not in database\n")
    # elif emp.upper() not in listEmployees and emp != '':
    #     window.infoConsole.append(f"Employee '{emp.upper()}' not in database\n")
    # else:
    #     if job and emp == '':
    #         window.infoConsole.append(f"Job '{job}' chosen\n")
    #         df = db[db['Job'] == job]
    #     elif job == '' and emp:
    #         window.infoConsole.append(f"Employee '{emp.upper()}' chosen\n")
    #         df = db[db['Employee'] == emp.upper()]
    #     elif job and emp:
    #         window.infoConsole.append(f"Job '{job}' and Employee '{emp.upper()}' chosen\n")
    #         df = db[(db['Employee'] == emp.upper()) & (db['Job'] == job)]
    #
    #     df.sort_values(['Employee', 'Date'], inplace=True)
    #     df['Job'] = df['Job'].astype('str')
    #     df.drop(columns=['Begin','End'], inplace=True)
    #     final = df.append(df.sum(numeric_only=True), ignore_index=True)

        # XLWINGS
        wb = xw.Book()  # this will create a new workbook
        sht = wb.sheets['Sheet1']
        sht.range('A1').options(index=False).value = final
        sht.range('A1').options(pd.DataFrame, expand='table').value
        totals_rangemaker = 'A'+str(final.shape[0]+1)+':'+'R'+str(final.shape[0]+1)
        sht.autofit('c')
        wb.sheets[0].range('A1:R1').api.Font.Bold = True
        wb.sheets[0].range(totals_rangemaker).api.Font.Bold = True
        return wb




if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())
