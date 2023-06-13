# pyinstaller   pyinstaller --onefile --noconsole --icon=swise.ico clinic_tool2.py

import sys
import pandas as pd
from PyQt5 import QtCore
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QPlainTextEdit, QTabWidget,
    QVBoxLayout, QListWidget, QLabel, QStackedWidget, QSizePolicy, QTextEdit, QSplashScreen, QPushButton
)
from PyQt5.QtCore import Qt, QRect
from PyQt5.QtGui import QIcon, QPixmap

class ClinicInfoTool(QWidget):
    def __init__(self, excel_file, parent=None):
        super().__init__()
        self.df = pd.read_excel(excel_file, sheet_name='Fac List', engine='openpyxl')
        self.init_ui()
        self.resize(800, 600)
        self.setWindowTitle('Clinic Info Tool')

    def run(self):
        splash = self.show_splash()  # Show the splash screen
        self.show()  # Show the main application window
        splash.finish(self)  # Hide the splash screen

    def show_splash(self):
        splash = QSplashScreen()
        splash.setWindowFlags(Qt.WindowStaysOnTopHint)
        label = QLabel("Loading, please wait...", splash)
        label.setStyleSheet("color: black; font-size: 16px;")
        splash.show()
        self.app.processEvents()
        return splash


    def init_ui(self):
        main_layout = QVBoxLayout()
        combined_layout = QHBoxLayout()

        # Search layout
        search_layout = QVBoxLayout()

        # Create a QHBoxLayout for input and buttons
        input_buttons_layout = QHBoxLayout()

        self.clinic_number_input = QLineEdit()
        self.clinic_number_input.setPlaceholderText('Type Clinic # and press Enter')
        self.clinic_number_input.returnPressed.connect(self.get_clinic_info)
        input_buttons_layout.addWidget(self.clinic_number_input)

        # Add the search button
        search_button = QPushButton('Search')
        search_button.clicked.connect(self.get_clinic_info)
        input_buttons_layout.addWidget(search_button)

        # Add the reset button
        reset_button = QPushButton('Reset')
        reset_button.clicked.connect(self.reset_to_defaults)
        input_buttons_layout.addWidget(reset_button)

        # Add the QHBoxLayout to the search_layout
        search_layout.addLayout(input_buttons_layout)

        self.result_text_edit = QTextEdit()
        self.result_text_edit.setReadOnly(True)
        search_layout.addWidget(self.result_text_edit)

        combined_layout.addLayout(search_layout)

        # Browse layout
        browse_layout = QVBoxLayout()

        self.groups_list = QListWidget()
        self.regions_list = QListWidget()
        self.areas_list = QListWidget()
        self.clinics_list = QListWidget()
        self.clinics_list.setSortingEnabled(True)
        self.clinics_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # Set the size policy
        self.clinics_list.setMinimumHeight(200)

        self.groups_list.itemClicked.connect(self.on_group_clicked)
        self.regions_list.itemClicked.connect(self.on_region_clicked)
        self.areas_list.itemClicked.connect(self.on_area_clicked)
        self.clinics_list.itemClicked.connect(self.on_clinic_clicked)  # Add this line

        self.update_groups()

        browse_layout.addWidget(QLabel('Groups'))
        browse_layout.addWidget(self.groups_list)
        browse_layout.addWidget(QLabel('Regions'))
        browse_layout.addWidget(self.regions_list)
        browse_layout.addWidget(QLabel('Areas'))
        browse_layout.addWidget(self.areas_list)
        browse_layout.addWidget(QLabel('Clinics'))
        browse_layout.addWidget(self.clinics_list)

        combined_layout.addLayout(browse_layout)

        main_layout.addLayout(combined_layout)
        self.setLayout(main_layout)

    def reset_to_defaults(self):
        self.clinic_number_input.clear()
        self.result_text_edit.clear()
        self.groups_list.clear()
        self.regions_list.clear()
        self.areas_list.clear()
        self.clinics_list.clear()
        self.update_groups()

    def _handle_nan(self, value):
        if pd.isna(value):
            return ""
        return value

    def get_clinic_info(self):
        clinic_number = self.clinic_number_input.text()

        if clinic_number == '':
            return  # Exit the method if the input is empty

        clinic_data = self.df.loc[self.df['Fac#'] == int(clinic_number)]

        if not clinic_data.empty:
            clinic_info = f"""
            <table>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Clinic Name:</b></td><td>{self._handle_nan(clinic_data['Clinic Name'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Street:</b></td><td>{self._handle_nan(clinic_data['Address'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>City, State, Zip:</b></td><td>{self._handle_nan(clinic_data['City'].values[0])}, {self._handle_nan(clinic_data['State'].values[0])} {self._handle_nan(clinic_data['Zip '].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Phone:</b></td><td>{self._handle_nan(clinic_data['Clinic PH / FX'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Clinic Manager:</b></td><td>{self._handle_nan(clinic_data['Clinic Manager'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Area:</b></td><td>{self._handle_nan(clinic_data['Area'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Area Manager:</b></td><td>{self._handle_nan(clinic_data['Area Team Lead (ATL)'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>DO:</b></td><td>{self._handle_nan(clinic_data['In-Center DO'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Region:</b></td><td>{self._handle_nan(clinic_data['REG'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>RVP:</b></td><td>{self._handle_nan(clinic_data['RVP'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Group:</b></td><td>{self._handle_nan(clinic_data['GRP'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Division:</b></td><td>{self._handle_nan(clinic_data['DIV'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>GVPO:</b></td><td>{self._handle_nan(clinic_data['GVP Name'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"></td><td></td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Additional Info:</b></td><td></td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>PAS Office Location:</b></td><td>{self._handle_nan(clinic_data['PAS Office Location'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>GVP/GM Assistant / Phone:</b></td><td>{self._handle_nan(clinic_data['GVP/GM Assistant / Phone'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>RVP Admin Assist / Phone:</b></td><td>{self._handle_nan(clinic_data['RVP Admin Assist / Phone'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Modalities Offered:</b></td><td>{self._handle_nan(clinic_data['Modalities Offered'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Clinic Details:</b></td><td>{self._handle_nan(clinic_data['Clinic Details'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Clip / Ph / Fx:</b></td><td>{self._handle_nan(clinic_data['Clip / Ph / Fx'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>PAS Supervisor:</b></td><td>{self._handle_nan(clinic_data['PAS Supervisor'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>PAS Supervisor Direct #:</b></td><td>{self._handle_nan(clinic_data['PAS Supervisor Direct #'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>PAS Team Lead:</b></td><td>{self._handle_nan(clinic_data['PAS Team Lead'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>PAS PICS:</b></td><td>{self._handle_nan(clinic_data['PAS PICS'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Medical Director:</b></td><td>{self._handle_nan(clinic_data['Medical Director'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Isolation?:</b></td><td>{self._handle_nan(clinic_data['Isolation?'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Escalation List (DO, RVP, HPSM, PAS TL, PAS Supervisor, etc):</b></td><td>{self._handle_nan(clinic_data['Escalation List (DO, RVP, HPSM, PAS TL, PAS Supervisor, etc)'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Clinical Quality Manager:</b></td><td>{self._handle_nan(clinic_data['Clinical Quality Manager'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Educators:</b></td><td>{self._handle_nan(clinic_data['Educators'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Revenue Center:</b></td><td>{self._handle_nan(clinic_data['Revenue Center'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>FC Supervisor:</b></td><td>{self._handle_nan(clinic_data['FC Supervisor'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Financial Coordinators:</b></td><td>{self._handle_nan(clinic_data['Financial Coordinators'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>VP of Marketing Development:</b></td><td>{self._handle_nan(clinic_data['VP of Marketing Development'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Dir of Marketing Development:</b></td><td>{self._handle_nan(clinic_data['Dir of Marketing Development'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Dir of HPS:</b></td><td>{self._handle_nan(clinic_data['Dir of HPS'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Dir. of Commercial Integrations:</b></td><td>{self._handle_nan(clinic_data['Dir. of Commercial Integrations'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>HPSM:</b></td><td>{self._handle_nan(clinic_data['HPSM'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>TOPS Coordinator:</b></td><td>{self._handle_nan(clinic_data['TOPS Coordinator'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Social Worker:</b></td><td>{self._handle_nan(clinic_data['Social Worker'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>In-Center DO Phone:</b></td><td>{self._handle_nan(clinic_data['In-Center DO Phone'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Home Therapy DO:</b></td><td>{self._handle_nan(clinic_data['Home Therapy DO'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Home Therapy DO Phone:</b></td><td>{self._handle_nan(clinic_data['Home Therapy DO Phone'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>GM Name:</b></td><td>{self._handle_nan(clinic_data['GM Name'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>GM Cell:</b></td><td>{self._handle_nan(clinic_data['GM Cell'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Commercial Extras:</b></td><td>{self._handle_nan(clinic_data['Commercial Extras'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>CIT Phone:</b></td><td>{self._handle_nan(clinic_data['CIT Phone'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>HPSM Phone:</b></td><td>{self._handle_nan(clinic_data['HPSM Phone'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>RFA Extras:</b></td><td>{self._handle_nan(clinic_data['RFA Extras'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>CVO Email for Blast:</b></td><td>{self._handle_nan(clinic_data['CVO Email for Blast'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>HT Group:</b></td><td>{self._handle_nan(clinic_data['HT Group'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Schedule Letter Extras:</b></td><td>{self._handle_nan(clinic_data['Schedule Letter Extras'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Sr. Manager SW Svcs:</b></td><td>{self._handle_nan(clinic_data['Sr. Manager SW Svcs'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>New Perm/NonFKC EIFs:</b></td><td>{self._handle_nan(clinic_data['New Perm/NonFKC EIFs'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Traveler EIFs:</b></td><td>{self._handle_nan(clinic_data['Traveler EIFs'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>CVO Group:</b></td><td>{self._handle_nan(clinic_data['CVO Group'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>OnBase Queue:</b></td><td>{self._handle_nan(clinic_data['OnBase Queue'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>TCU:</b></td><td>{self._handle_nan(clinic_data['TCU'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Transport Program:</b></td><td>{self._handle_nan(clinic_data['Transport Program'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>BC Case Manager:</b></td><td>{self._handle_nan(clinic_data['BC Case Manager'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>TCU Days/Week:</b></td><td>{self._handle_nan(clinic_data['TCU Days/Week'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>KCA:</b></td><td>{self._handle_nan(clinic_data['KCA'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Dietitian:</b></td><td>{self._handle_nan(clinic_data['Dietitian'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Sr. Manager Clinical Quality:</b></td><td>{self._handle_nan(clinic_data['Sr. Manager Clinical Quality'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Sr. Manager Nutrition Svcs:</b></td><td>{self._handle_nan(clinic_data['Sr. Manager Nutrition Svcs'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Manager Nutrition Svcs:</b></td><td>{self._handle_nan(clinic_data['Manager Nutrition Svcs'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Manager SW Svcs:</b></td><td>{self._handle_nan(clinic_data['Manager SW Svcs'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Sr. Manager Clinical Education:</b></td><td>{self._handle_nan(clinic_data['Sr. Manager Clinical Education'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Clinical Educator:</b></td><td>{self._handle_nan(clinic_data['Clinical Educator'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Clinic County:</b></td><td>{self._handle_nan(clinic_data['Clinic County'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>eCC Instance:</b></td><td>{self._handle_nan(clinic_data['eCC Instance'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>CVO Special Note:</b></td><td>{self._handle_nan(clinic_data['CVO Special Note'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>PAS Manager:</b></td><td>{self._handle_nan(clinic_data['PAS Manager'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>FAS Leadership:</b></td><td>{self._handle_nan(clinic_data['FAS Leadership'].values[0])}</td></tr>
                <tr><td style="text-align: right; padding-right: 10px;"><b>Senior HPSM:</b></td><td>{self._handle_nan(clinic_data['Senior HPSM'].values[0])}</td></tr>
            </table>

            """
            self.result_text_edit.setHtml(clinic_info)

            # Find and select the corresponding Group, Region, and Area
            group_name = clinic_data['GRP'].values[0].lower().strip()
            for i in range(self.groups_list.count()):
                item_text = self.groups_list.item(i).text().lower()
                if group_name in item_text:
                    self.groups_list.setCurrentRow(i)
                    self.on_group_clicked(self.groups_list.item(i))
                    break

            region_name = clinic_data['REG'].values[0].lower().strip()
            for i in range(self.regions_list.count()):
                item_text = self.regions_list.item(i).text().lower()
                if region_name in item_text:
                    self.regions_list.setCurrentRow(i)
                    self.on_region_clicked(self.regions_list.item(i))
                    break

            area_name = clinic_data['Area'].values[0].lower().strip()
            for i in range(self.areas_list.count()):
                item_text = self.areas_list.item(i).text().lower()
                if area_name in item_text:
                    self.areas_list.setCurrentRow(i)
                    self.on_area_clicked(self.areas_list.item(i))
                    break

        else:
            self.result_text_edit.setPlainText('Clinic not found.')

    def update_groups(self):
        self.groups_list.clear()
        group_names = self.df['GRP'].str.lower().str.strip().unique()
        group_names = [g for g in group_names if g and isinstance(g, str)]  # Filter out empty strings and non-string values
        group_names.sort()
        self.groups_list.addItem("All Groups")
        for group_name in group_names:
            gvp_name = self._handle_nan(self.df[self.df['GRP'].str.lower().str.strip() == group_name]['GVP Name'].values[0])
            display_text = f"{group_name.title()} (GVP: {gvp_name})"
            self.groups_list.addItem(display_text)


    def update_regions(self):
        self.regions_list.clear()
        region_names = self.df['REG'].apply(lambda x: str(x).lower().strip() if isinstance(x, str) else '')
        region_names = region_names[region_names != ''].unique()
        region_names.sort()
        self.regions_list.addItem("All Regions")
        for region_name in region_names:
            rvp = self._handle_nan(self.df[self.df['REG'].str.lower().str.strip() == region_name]['RVP'].iloc[0])
            display_text = f"{region_name.title()} (RVP: {rvp})"
            self.regions_list.addItem(display_text)

    def update_areas(self):
        self.areas_list.clear()
        area_names = self.df['Area'].str.lower().str.strip().unique()
        area_names = [a for a in area_names if a]  # Filter out empty strings
        area_names.sort()
        self.areas_list.addItem("All Areas")
        for area_name in area_names:
            do = self._handle_nan(self.df[self.df['Area'].str.lower().str.strip() == area_name]['In-Center DO'].iloc[0])
            display_text = f"{area_name.title()} (DO: {do})"
            self.areas_list.addItem(display_text)


    def on_group_clicked(self, item):
        group_name = item.text().lower().split(' (gvp:', 1)[0]
        self.regions_list.clear()
        self.areas_list.clear()
        self.clinics_list.clear()
        if group_name.lower() == "all groups":
            self.update_regions()
            self.update_areas()
            self.update_clinics()
        else:
            regions = (
                self.df[self.df['GRP'].str.lower().str.strip() == group_name]
            )
            region_names = regions['REG'].str.lower().str.strip().unique()
            region_names = [r for r in region_names if r]  # Filter out empty strings
            region_names.sort()
            self.regions_list.addItem("All Regions")
            for region_name in region_names:
                rvp = regions[regions['REG'].str.lower().str.strip() == region_name]['RVP'].iloc[0]
                display_text = f"{region_name.title()} (RVP: {rvp})"
                self.regions_list.addItem(display_text)

    def on_region_clicked(self, item):
        region_name = item.text().lower().split(' (rvp:', 1)[0]  # Remove the RVP part from the text
        self.areas_list.clear()
        self.clinics_list.clear()
        if region_name.lower() == "all regions":
            self.update_areas()
        else:
            areas = (
                self.df[self.df['REG'].str.lower().str.strip() == region_name]
            )
            area_names = areas['Area'].str.lower().str.strip().unique()
            area_names = [a for a in area_names if a]  # Filter out empty strings
            area_names.sort()
            self.areas_list.addItem("All Areas")
            for area_name in area_names:
                do = areas[areas['Area'].str.lower().str.strip() == area_name]['In-Center DO'].iloc[0]
                do = self._handle_nan(do)
                display_text = f"{area_name.title()} (DO: {do})"
                self.areas_list.addItem(display_text)

    def on_area_clicked(self, item):
        area_name = item.text().split(' (do:', 1)[0]  # Remove the DO part from the text
        self.update_clinics(area_name)

    def on_clinic_clicked(self, item):
        clinic_number = item.text().split(' -', 1)[0]  # Extract the clinic number from the text
        self.clinic_number_input.setText(clinic_number)  # Set the clinic number in the input field
        self.get_clinic_info()  # Call the method to display the clinic details

    def update_clinics(self, area_name: str):
        self.clinics_list.clear()
        area_name = area_name.lower().split(' (do:', 1)[0]
        if area_name == "all areas":
            clinics = self.df['Fac#'].unique()
        else:
            clinics = self.df[self.df['Area'].str.lower().str.strip() == area_name]['Fac#'].unique()

        clinics = list(map(str, clinics))  # Convert all elements to strings
        clinics = [c for c in clinics if c]  # Filter out empty strings
        clinics.sort()

        for clinic in clinics:
            if clinic.isdigit():  # Ignore non-numeric clinic numbers
                clinic_data = self.df[self.df['Fac#'] == int(clinic)].iloc[0]
                clinic_name = clinic_data['Clinic Name']
                cm = self._handle_nan(clinic_data['Clinic Manager'])
                display_text = f"{clinic} - {clinic_name} (CM: {cm})"
                self.clinics_list.addItem(display_text)


if __name__ == '__main__':
    app = QApplication(sys.argv)

    # Create and show the splash screen
    splash = QSplashScreen()
    splash.setWindowFlags(Qt.WindowStaysOnTopHint)

    splash_image = QPixmap("\\\\corpfs01\\fmcna-shared\\OPEX\\chrome\\data\\SWICO.png")  # Replace with the path to your image
    splash = QSplashScreen(splash_image)
    label = QLabel("Loading, please wait...", splash)
    label.setStyleSheet("""
        color: white;
        font-size: 18px;
        font-weight: bold;
        background-color: rgba(0, 0, 0, 100);
        padding: 10px;
        border-radius: 5px;
    """)

    label_width = label.sizeHint().width()
    label_height = label.sizeHint().height()
    image_width = splash_image.width()
    image_height = splash_image.height()

    label.setGeometry(
        QRect((image_width - label_width) // 2, (image_height - label_height) // 2, label_width, label_height)
    )

    splash.show()
    app.processEvents()

    excel_file = r'\\corpfs01\fmcna-shared\OPEX\chrome\data\FMC Clinic Info Tool2-SW.xlsm'
    app.setWindowIcon(QIcon('\\corpfs01\fmcna-shared\OPEX\chrome\data\swise.ico'))  # Set the application icon

    clinic_info_tool = ClinicInfoTool(excel_file)
    clinic_info_tool.app = app
    clinic_info_tool.show()  # Show the main application window
    splash.finish(clinic_info_tool)  # Hide the splash screen
    clinic_info_tool.setWindowIcon(QIcon('swise.ico'))  # Set the window icon

    sys.exit(app.exec_())
