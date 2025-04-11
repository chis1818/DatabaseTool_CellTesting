
# Tkinter Google Sheets Row Updater for CPP FSAEE 25

**This tool was developed for CPP FSAEE 25, to make cell testing and categorization easier.**

## Project Overview

**Tkinter Google Sheets Row Updater** is a Python-based desktop GUI tool that allows you to search for and update rows in a Google Sheets document by using a unique identifier from the first column (e.g., a serial number). It features:
- A search field (labeled dynamically with the text in cell A1 of your spreadsheet).
- Editable entry fields that display data from the row.
- The ability to update only non-blank fields so that existing data is not overridden with blanks.
- A feedback label that displays the row number that was found or updated.
- A window title that reflects the Google Sheet’s title.
- An optional custom icon (via an `.ico` file).

## Features

- **Search by First Column:**  
  Enter a unique identifier (e.g. serial number) that exists in the first column. The tool will search for that value and populate the rest of the row into editable fields.
  
- **Selective Field Update:**  
  Only fields that are changed (non-empty in the input) are updated. Blank fields do not overwrite the existing cell values in the Google Sheet.

- **Dynamic Labels & Window Title:**  
  - The label for the search input is taken from cell A1 (e.g., “Serial Number” or “ID”).
  - The application window title automatically displays the name of your Google Sheet.
  
- **Feedback Mechanism:**  
  A status label informs you if the requested row was found and indicates which row was updated.

- **Optional Custom Icon:**  
  If an `.ico` file is present in the project folder, it will be used as the window icon. Otherwise, the default Tkinter icon is used.

## Folder & File Setup

Place **all necessary files in the same folder**:

```
/your_project_folder
  ├── update_row_tool.py      # Main Python script
  ├── cell-database-456510-a79db50a2ddd.json   # Service account JSON file (required)
  └── CPPFSAEE.ico            # Optional custom window icon
```

### Important Files
- **Service Account JSON:** The Google credentials file (downloaded from Google Cloud).
- **Icon File (Optional):** An `.ico` file to customize the window icon.

## Google Cloud Setup Instructions

Before running the tool, you must set up access to your Google Sheets:

1. **Create a Google Cloud Project:**  
   - Go to the [Google Cloud Console](https://console.cloud.google.com/) and create a new project (or use an existing project).

2. **Enable APIs:**  
   - Enable the **Google Sheets API**.
   - Enable the **Google Drive API**.

3. **Create a Service Account:**  
   - Navigate to **APIs & Services > Credentials**.
   - Click **Create Credentials > Service Account**.
   - Provide a name (e.g., "SheetsUpdater Service Account") and assign an appropriate role (e.g., "Editor").

4. **Generate a JSON Key:**  
   - After the service account is created, generate a key by selecting JSON as the key type.
   - Download the JSON file and rename it if necessary (e.g., `cell-database-456510-a79db50a2ddd.json`).
   - Place this file in the same folder as the Python script.

5. **Share Your Google Sheet:**  
   - Open your target Google Sheet.
   - Click on **Share** and add the **service account email** (found inside the JSON file under `client_email`).
   - Grant the service account **edit access**.

## How to Run the Tool

1. **Install Dependencies:**  
   Ensure you have Python 3.6+ installed. Then open a terminal (or command prompt) in the project folder and run:
   ```bash
   pip install gspread oauth2client
   ```

2. **Launch the Application:**  
   - You can launch the tool by double-clicking the Python script (if Python is associated with `.py` files on your system) **or** run it from the terminal with:
     ```bash
     python update_row_tool.py
     ```
     
3. **Using the GUI:**  
   - The search field is labeled according to the text in cell A1 of your spreadsheet.
   - Enter the first-column value (e.g., a serial number) and click **Search**.  
     The tool will populate the fields with the current values in the row and update a status label (e.g., "Serial 'XYZ' found on row: 5").
   - Modify the necessary fields. Leave any field blank if you don't want it updated.
   - Click **Update** to apply the changes. The status label will then indicate which row was updated (e.g., "Serial 'XYZ' updated on row: 5").

## Customization Notes

- **Configuration Variables:**  
  At the top of the Python script, you can easily change:
  - `SERVICE_ACCOUNT_FILENAME` (the name of your service account JSON file)
  - `ICON_FILENAME` (the icon file name, if used)
  - `SPREADSHEET_URL` (the URL of your target Google Sheet)
  - `WORKSHEET_NAME` (the tab name inside your Google Sheet)

- **Search Label & Window Title:**  
  The script automatically uses cell A1 for the search field label and the Google Sheet’s name as the window title.

- **Icon Customization:**  
  The provided `.ico` file is optional. If it is not present, the default Tkinter icon is used.

## Requirements

- **Python 3.6+**
- **Internet Connection** (required for accessing the Google Sheets API)
- **Google Service Account** with proper credentials and spreadsheet access
- **Installed Dependencies:** `gspread`, `oauth2client`
