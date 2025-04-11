import os
import tkinter as tk
import tkinter.messagebox as mb
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ================
# USER CONFIGURATION
# ================
# Replace with your service account JSON file name (must be in the same folder)
SERVICE_ACCOUNT_FILENAME = "service_account.json"

# (Optional) Replace with your custom icon filename; if missing, default icon is used.
ICON_FILENAME = "app.ico"

# Replace with your Google Sheet URL
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit"

# Replace with the name of the worksheet (tab) in your Google Sheet
WORKSHEET_NAME = "Sheet1"
# -------------------------------------
  
# Build absolute paths for files assumed to be in the same directory as this script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, SERVICE_ACCOUNT_FILENAME)
ICON_PATH = os.path.join(BASE_DIR, ICON_FILENAME)

def authorize_gspread():
    """Authorize gspread using a service account JSON file."""
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
    return gspread.authorize(creds)

class SerialUpdaterApp(tk.Tk):
    """
    Tkinter GUI Application for updating a Google Sheet row:
    
    1. Reads headers from the first row of the Sheet.
    2. Displays editable fields (for columns 2 onward) and uses cell A1 as the search label.
    3. Searches for a row by the first column value (e.g., a serial number).
    4. Updates only non-blank fields in that row.
    5. Sets the window title as the Google Sheet’s title.
    6. Optionally uses a custom window icon.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Attempt to load the custom window icon.
        try:
            self.iconbitmap(ICON_PATH)
        except Exception as e:
            print(f"Could not load icon '{ICON_PATH}': {e}")

        # Authorize and open the Google Sheet.
        self.gc = authorize_gspread()
        self.sh = self.gc.open_by_url(SPREADSHEET_URL)
        self.ws = self.sh.worksheet(WORKSHEET_NAME)

        # Set the window title to the sheet's title.
        self.title(self.sh.title)

        # Read the header row (row 1) of the sheet.
        self.headers = self.ws.row_values(1)
        if not self.headers:
            mb.showerror("Error", "No headers found in row 1 of the sheet.")
            self.destroy()
            return

        # Set up the main UI components.
        self._create_widgets()
        
        # Create entry widgets for fields (columns 2 and onward).
        self.entry_widgets = []
        for idx, header in enumerate(self.headers[1:], start=1):
            lbl = tk.Label(self.frm_fields, text=header + ":")
            lbl.grid(row=idx-1, column=0, padx=5, pady=2, sticky="e")
            ent = tk.Entry(self.frm_fields, width=30)
            ent.grid(row=idx-1, column=1, padx=5, pady=2)
            self.entry_widgets.append(ent)

    def _create_widgets(self):
        """Set up layout: data fields, a status label, and a bottom frame with search/update controls."""
        # Frame for data fields.
        self.frm_fields = tk.Frame(self)
        self.frm_fields.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(10, 5))

        # Status label.
        self.var_info = tk.StringVar(value="")
        self.lbl_info = tk.Label(self, textvariable=self.var_info)
        self.lbl_info.pack(side=tk.TOP, pady=(5, 0))

        # Bottom frame for search and update controls.
        frm_bottom = tk.Frame(self)
        frm_bottom.pack(side=tk.BOTTOM, fill=tk.X, pady=5)

        # Use the first header cell (A1) as the label for the search field.
        a1_text = self.headers[0] if self.headers else "Search Value"
        lbl_sn = tk.Label(frm_bottom, text=f"{a1_text}:")
        lbl_sn.pack(side=tk.LEFT, padx=(10, 5))

        # Entry widget for searching.
        self.var_serial = tk.StringVar()
        ent_sn = tk.Entry(frm_bottom, textvariable=self.var_serial, width=25)
        ent_sn.pack(side=tk.LEFT, padx=5)

        # Search button.
        btn_search = tk.Button(frm_bottom, text="Search", command=self.on_search)
        btn_search.pack(side=tk.LEFT, padx=5)

        # Update button.
        btn_update = tk.Button(frm_bottom, text="Update", command=self.on_update)
        btn_update.pack(side=tk.LEFT, padx=5)

    def on_search(self):
        """Search for a row by the value in column A and populate entry fields; update status label."""
        serial = self.var_serial.get().strip()
        if not serial:
            mb.showerror("Error", f"Please enter a {self.headers[0]} to search.")
            return

        try:
            cell = self.ws.find(serial, in_column=1)
        except gspread.exceptions.CellNotFound:
            cell = None

        if not cell:
            mb.showerror("Not Found", f"{self.headers[0]} '{serial}' not found.")
            return

        target_row = cell.row
        row_values = self.ws.row_values(target_row)
        while len(row_values) < len(self.headers):
            row_values.append("")

        for i, ent in enumerate(self.entry_widgets, start=1):
            ent.delete(0, tk.END)
            ent.insert(0, row_values[i])

        self.var_info.set(f"{self.headers[0]} '{serial}' found on row: {target_row}")

    def on_update(self):
        """Update the Google Sheet row with the non-blank field values and update the status label."""
        serial = self.var_serial.get().strip()
        if not serial:
            mb.showerror("Error", f"Please enter a {self.headers[0]} to update.")
            return

        try:
            cell = self.ws.find(serial, in_column=1)
        except gspread.exceptions.CellNotFound:
            cell = None

        if not cell:
            mb.showerror("Not Found", f"{self.headers[0]} '{serial}' not found.")
            return

        target_row = cell.row
        old_row_values = self.ws.row_values(target_row)
        while len(old_row_values) < len(self.headers):
            old_row_values.append("")

        new_values = []
        for i, ent in enumerate(self.entry_widgets):
            typed_value = ent.get().strip()
            if typed_value:
                new_values.append(typed_value)
            else:
                new_values.append(old_row_values[i+1] if i+1 < len(old_row_values) else "")

        last_col_index = len(self.headers)
        first_col_letter = chr(64 + 2)  # Column B
        last_col_letter = chr(64 + last_col_index)
        update_range = f"{first_col_letter}{target_row}:{last_col_letter}{target_row}"
        data_2d = [new_values]

        try:
            self.ws.update(update_range, data_2d, value_input_option="USER_ENTERED")
            self.var_info.set(f"{self.headers[0]} '{serial}' updated on row: {target_row}")
        except Exception as e:
            mb.showerror("Error Updating", str(e))

def main():
    app = SerialUpdaterApp()
    app.mainloop()

if __name__ == "__main__":
    main()
