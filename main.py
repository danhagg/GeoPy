import csv
import folium
import io
import math
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
        self.queryButton = QPushButton("Query Cocorahs Database")
        self.queryButton.clicked.connect(self.query_on_click)

        # START DATES
        self.currentDateEdit = QDateEdit()
        self.currentDateEdit.setDisplayFormat('dd MMM yyyy')
        self.currentDateEdit.setDate(self.calendar.selectedDate())
        self.currentDateEdit.setDateRange(self.calendar.minimumDate(),
                self.calendar.maximumDate())
        self.currentDateLabel = QLabel("&Date:")
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
        self.textboxLongLabel = QLabel("Target Longitude:")

        # LATITUDE
        self.textboxLat = QLineEdit(self)
        self.textboxLatLabel = QLabel("Target Latitude:")

        # STATES
        self.textboxStates = QLineEdit(self)
        self.textboxStatesLabel = QLabel("State(s): e.g., 'CO,NE,WY'")

        dateBoxLayout = QGridLayout()
        dateBoxLayout.addWidget(self.currentDateLabel, 1, 0)
        dateBoxLayout.addWidget(self.currentDateEdit, 1, 1, 1, 2)
        # dateBoxLayout.addWidget(self.maximumDateLabel, 2, 0)
        # dateBoxLayout.addWidget(self.maximumDateEdit, 2, 1, 1, 2)
        dateBoxLayout.addWidget(self.textboxLatLabel, 3, 0)
        dateBoxLayout.addWidget(self.textboxLat, 3, 1, 1, 2)
        dateBoxLayout.addWidget(self.textboxLongLabel, 4, 0)
        dateBoxLayout.addWidget(self.textboxLong, 4, 1, 1, 2)
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
        states = self.textboxStates.text()
        date_range = list(range(startDate.month(), endDate.month() + 1))
        dateString = startDate.toString('MM/dd/yyyy')
        query(long=long, lat=lat, dateString=dateString, states=states)

#http://data.cocorahs.org/export/exportreports.aspx?ReportType=hail&Format=csv&State=TX&ReportDateType=reportdate&Date=8/14/2018

def query(long, lat, dateString, states):
    if long == '' or lat == '':
        window.infoConsole.append("Empty Target Coordinates")
        return
    elif states == '':
        window.infoConsole.append("Empty State")
        return
    else:
        CSV_URL = 'http://data.cocorahs.org/export/exportreports.aspx'

        payload = {'ReportType': 'hail', 'Format': 'csv', 'State': states, 'ReportDateType': 'reportdate', 'Date': dateString}

        with requests.Session() as s:
            data = s.get(CSV_URL, params=payload)


    df = pd.read_csv(io.StringIO(data.text), usecols=['StationName', 'Latitude', 'Longitude', 'AverageSize', 'DateTimeStamp'])
    if df.empty == True:
        window.infoConsole.append("Cocorahs database has no data for that combination of state(s) and date(s)")
        return
    else:
        df['TargetLatitude'] = lat
        df['TargetLatitude'] = df['TargetLatitude'].astype('float')
        df['TargetLongitude'] = long
        df['TargetLongitude'] = df['TargetLongitude'].astype('float')
        df['Distance'] = df[['Latitude', 'Longitude','TargetLatitude','TargetLongitude']].apply(lambda x: distance(*x), axis=1)
        final = df


        map_osm = folium.Map(location=[float(lat), float(long)], tiles="Stamen Toner", zoom_start=8)

        locations = df[['Latitude', 'Longitude']]
        locationlist = locations.values.tolist()
        for point in range(0, len(locationlist)):
            folium.Marker(locationlist[point], popup=df['StationName'][point]).add_to(map_osm)

            # reduce this to one marker
        df.apply(lambda row:folium.CircleMarker(location=[row["TargetLatitude"], row["TargetLongitude"]],
                                                      radius=10, fill_color='red')
                                                     .add_to(map_osm), axis=1)

        map_osm.save('map_osm.html')

        # XLWINGS
        wb = xw.Book()  # this will create a new workbook
        sht = wb.sheets['Sheet1']
        sht.range('A1').options(index=False).value = final
        sht.range('A1').options(pd.DataFrame, expand='table').value
        sht.autofit('c')
        wb.sheets[0].range('A1:R1').api.Font.Bold = True
        return wb



def distance(stationLatitude, stationLongitude, targetLatitude, targetLongitude):
    lat1, lon1 = stationLatitude, stationLongitude
    lat2, lon2 = float(targetLatitude), float(targetLongitude)
    km_radius = 6371 # km
    mile_radius = 3959 # radius of the great circle in miles

    dlat = math.radians(lat2-lat1)
    dlon = math.radians(lon2-lon1)
    a = math.sin(dlat/2) * math.sin(dlat/2) + math.cos(math.radians(lat1)) \
        * math.cos(math.radians(lat2)) * math.sin(dlon/2) * math.sin(dlon/2)
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))
    d = mile_radius * c

    return d


def download_file(url, filename):
    ''' Downloads file from the url and save it as filename '''
    # check if file already exists
    if not os.path.isfile(filename):
        print('Downloading File')
        response = requests.get(url)
        # Check if the response is ok (200)
        if response.status_code == 200:
            # Open file and write the content
            with open(filename, 'wb') as file:
                # A chunk of 128 bytes
                for chunk in response:
                    file.write(chunk)
    else:
        print('File exists')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())
