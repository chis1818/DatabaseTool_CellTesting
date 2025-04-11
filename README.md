# DatabaseTool_CellTesting
Python script that access's google sheets using google cloud api and allows users to update infoamtion about a specific serial number. Created for use to organize and manage CPPFSAEE25 21700 Cells
`
Google Sheets Row Updater (Tkinter GUI Tool)
Google Sheets Row Updater is a Python-based desktop GUI tool (using Tkinter) for updating specific rows in a Google Sheet by searching for a serial number (or any unique identifier in the first column). It connects to the Google Sheets API via a service account and provides a simple interface to edit row data. The tool only updates cells that have new, non-blank values entered, leaving all other cells unchanged to avoid overwriting existing data.
Project Overview
This project offers an easy way to modify entries in a Google Sheets spreadsheet through a local GUI application. Instead of manually finding a row in Google Sheets, you can use this tool to search for a row by its serial number (or first-column label) and update certain fields quickly. The GUI displays the current values of the row and allows you to edit them. When you submit an update, only the fields you filled in (non-blank inputs) will be written to the Google Sheet, so no data will be erased unintentionally. The app provides immediate feedback on which row was found and updated, and it dynamically uses your spreadsheet’s name in the window title for clarity.
Features
Search by Serial Number (First Column): Input a serial number (or any value from the first column of the sheet) to locate the corresponding row. The first cell (A1) of your sheet is used as the label for this search field (e.g. if A1 is "Asset ID", the search field will be labeled "Asset ID").
Edit Row Fields: Once a matching row is found, the tool displays the rest of that row’s values in editable fields. You can modify one or more of these values in the GUI.
Update Non-Blank Entries Only: When you click the update button, the program will send updates only for fields you have changed (i.e. fields you did not leave blank). This ensures that no cell will be overwritten with a blank value accidentally – existing data is left untouched if its field is left blank in the form.
Feedback Label: The GUI includes a status/feedback label that confirms the operation – for example, it may show a message like “Row 5 updated successfully” or “No matching serial number found.” This helps confirm which row was found and updated.
Dynamic Window Title: The application’s window title automatically shows the name of your Google Sheet. This makes it clear which spreadsheet you are connected to, especially if you work with multiple sheets.
Custom Icon (Optional): You can provide a custom icon for the application window by placing an .ico file in the application folder. If an icon file is present, the GUI will use it for the window; otherwise, the default Tkinter icon is used. This is useful for branding or easy identification of the tool on your desktop or taskbar.
Folder and File Setup
To set up the tool, make sure all the necessary files are located in the same folder (project directory). These include:
The main Python script (the GUI application’s .py file).
The Google service account credentials JSON file (downloaded from Google Cloud – see Setup instructions below). Place this JSON file in the same folder as the script for the tool to find it (you may rename it to something like credentials.json for convenience)​
codearmo.com
.
(Optional) An icon file for the app window (e.g. app.ico). If provided in this folder, the tool will use it as the window icon.
Keeping these files together in one directory ensures the program can locate the credentials and icon when running.
Google Cloud Setup Instructions
Before running the tool, you need to configure a Google Cloud project to allow access to Google Sheets. Follow these steps to set up the Google Sheets API credentials:
Create a Google Cloud Project: Go to the Google Cloud Console and create a new project (or use an existing project) for this tool. This project will hold your API settings and credentials.
Enable the Google Sheets and Drive APIs: In the Cloud Console, navigate to APIs & Services > Library. Enable the Google Sheets API for your project. Also enable the Google Drive API, which is needed to allow file access permissions for the service account​
dev.to
. (Enabling the Drive API ensures the application can access the spreadsheet file in Google Drive.)
Create a Service Account: In the Cloud Console, go to APIs & Services > Credentials and click Create Credentials > Service Account. Give the service account a name (e.g. "SheetsUpdater Service Account"). During this process, you can assign it a role; for example, choose Project > Editor to allow the account to edit Google Sheets on your behalf​
dev.to
. Complete the service account creation steps.
Generate a Key (JSON File): Still on the Service Account credentials page, create a new key for the service account. Choose the JSON key type and download the key file​
dev.to
. This file contains the service account’s credentials. Save this JSON file to the same folder as the Python script (as mentioned above). Keep it secure, as it provides access to your Google account project.
Share the Google Sheet with the Service Account: Open the target Google Sheet that you want to update. Click the Share button in the sheet, and add the service account’s email (you can find this email in the JSON key file under the client_email field). Give this email address edit access to the spreadsheet​
denisluiz.medium.com
. (The service account needs permission to edit the sheet; sharing the document with the service account email is essential​
stackoverflow.com
.) Note: If the service account is not shared on the sheet, the tool will not be able to find or update any data on that spreadsheet. Make sure the sharing settings are correct before running the application.
With these steps completed, your Google Cloud project is set up and the service account JSON is in place. The application will use these credentials to authenticate and connect to your Google Sheet.
Running the Script
Once you have all files in place and the Google API configured, follow these steps to run the tool:
Install Dependencies: Ensure you have Python 3.6 or above installed. Install the required Python libraries by running:
bash
Copy
Edit
pip install gspread oauth2client
This will install the gspread library (for interacting with Google Sheets) and oauth2client (for Google OAuth2 authentication) along with any other needed dependencies​
codearmo.com
. You can install these via a terminal/command prompt. (If you already have these or if your project uses a requirements.txt, make sure these packages are included.)
Launch the Application: After installing dependencies, you can start the GUI tool. Simply double-click the Python script file (e.g. update_row_tool.py) in your file explorer, and the Tkinter GUI window will open. Alternatively, you can run the script from a terminal with the command:
bash
Copy
Edit
python update_row_tool.py
Once launched, the application window should appear. Enter a serial number (or the appropriate first-column value) in the search field and press the search/update button. The status label will indicate if a row was found, and the fields will populate with that row’s current data for you to edit. After making changes, click the update button to submit the changes to the Google Sheet. The status label will update to confirm the row number that was updated.
Customization Notes
Search Field Label: The tool uses the content of cell A1 in your Google Sheet as the label for the search field. For example, if cell A1 of the sheet contains “Serial Number” or “ID”, the application will display that text next to the search input box. This makes the UI adaptable to different sheets (you know exactly what to enter as the search key).
Window Title: The title of the application window is set to the name of your Google Spreadsheet. If your Google Sheet is named “Inventory 2025”, the GUI’s title bar will show “Inventory 2025” (along with maybe the app name or description). This helps verify at a glance which sheet you are connected to, especially if you have multiple sheets or instances open.
.ico File for Icon: If an .ico icon file is present in the application’s folder, the tool will use it for the window icon. You can use this to give the app a custom icon (for example, a company logo or a relevant symbol). To use it, place your .ico file in the same directory as the script and name it accordingly (check the script for the expected icon filename, e.g., app.ico). If no icon file is found, the application will simply use the default Tkinter icon. This feature is optional and mainly for cosmetic/custom branding purposes.
Requirements
Make sure your environment meets the following requirements:
Python 3.6+ – The script is written for Python 3 and uses Tkinter, which is included with standard Python installations. It should run on Python 3.6 or newer.
Internet Connection – An active internet connection is required for the tool to access the Google Sheets API. Without internet, the app cannot fetch or update data on Google’s servers.
Google Sheet Access – Access to the target Google Sheet via the service account. This means you have completed the Google Cloud setup and shared the spreadsheet with the service account’s email. The Google account used for setup should have the Google Sheets and Drive APIs enabled as described, and the service account JSON must be present in the app folder.
With all the above in place, you can use the Google Sheets Row Updater tool to conveniently search and update rows in your spreadsheet. Enjoy streamlined data entry and editing without opening your browser! Note: Always keep your service account JSON credentials secure and avoid sharing them. Anyone with this file can potentially access your Google Sheets data. Only use this tool with sheets and data you trust and have permission to modify.
