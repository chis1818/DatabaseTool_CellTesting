# Tkinter Google Sheets Row Updater

**Tkinter Google Sheets Row Updater** is a Python-based desktop GUI tool that lets you search for a particular row in a Google Sheet by its first-column value (often a serial number) and update only the fields you modify. It is especially useful for quick data-entry edits without needing to open Google Sheets in the browser.

## Project Overview

- **Purpose**: Provide a local GUI tool to search and update rows in a Google Sheet.
- **Updates Non-Blank Fields**: Only overwrites data for fields that you explicitly edit. Fields left blank in the GUI remain unchanged in the Sheet.
- **Feedback Label**: Shows the row found during search, or which row was updated after making changes.
- **Window Title**: Automatically displays the spreadsheet’s name for clarity.
- **Optional Icon**: You may provide an `.ico` file to brand/customize the window icon.

## Features

1. **Search by First Column**  
   Enter a value (e.g. a serial number) that exists in the first column of your sheet. If found, the tool will show the current data in the rest of that row.
   
2. **Editable Fields**  
   Fields from columns 2 onward are editable in the GUI. You can change multiple fields at once.

3. **Selective Update**  
   After editing, clicking **Update** sends only non-empty fields to the Google Sheet. Blank fields are ignored so that no existing data is wiped out unintentionally.

4. **Live Status Label**  
   A status label at the bottom of the GUI tells you when a row is found or updated, showing which row number was affected.

5. **Automatic Labels**  
   The label for the search box uses whatever text is in cell A1 (for instance, “Serial Number” or “ID”). The application window title is set to the spreadsheet’s title (from your Google Sheet).

6. **Optional Custom Icon**  
   Placing an `.ico` file in the same folder will use that icon for your application window, otherwise the default Tkinter icon is shown.

## Folder & File Setup

Put all necessary files into one folder:
- The **main Python script** (e.g. `update_row_tool.py`).
- The **service account JSON** file (from Google Cloud) for your credentials (e.g. `service_account_key.json`).
- *(Optional)* The **.ico icon file** to change the app window icon (e.g. `app.ico`).

