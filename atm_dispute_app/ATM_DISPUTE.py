# ...existing code...
from constants import BGL_DR, SYS_DR, CUSTOMER_DR, BGL_CR, SYS_CR, CUSTOMER_CR, checkDigitConstantArray, MAX_CR_DR_AMOUNT_ALLOWED
import error_code as ec

class InputFileValidationRules:
    # ...existing code...

    def checkAccountNumber(self, record, fileobj):
        acc_no = record.getAccountNumber()
        if acc_no is None:
            record.setError(True)
            record.setErrorCode(ec.ACCOUNT_NUMERIC_ERR)
            return

        if acc_no == 0:
            record.setError(True)
            record.setErrorCode(ec.ACCOUNT_NUMBER_ZERO_ERR)
            return

        account_str = str(acc_no)
        account_length = len(account_str)

        if not (11 <= account_length <= 13):
            record.setError(True)
            record.setErrorCode(ec.ACCOUNT_NUMBER_CHARS_ERR)
            return

        last_digit = acc_no % 10
        base_no = acc_no // 10
        calculated_check_digit = self.getCheckDigitNumber(base_no)

        if last_digit != calculated_check_digit:
            record.setError(True)
            record.setErrorCode(ec.ACCOUNT_CHK_DIGIT_ERR)

    # ...existing code...

    def getCheckDigitNumber(self, account_number):
        """Exact Java algorithm port"""
        Macno = int(account_number)
        JE = 15
        Mchkdigit = 0

        while Macno > 0 and JE >= 0:
            iLastDigit = Macno % 10
            if iLastDigit > 0:
                Mdigit = iLastDigit - 1
                indexI = JE
                indexJ = Mdigit
                if 0 <= indexI < len(checkDigitConstantArray) and 0 <= indexJ < len(checkDigitConstantArray[0]):
                    Mchkdigit += checkDigitConstantArray[indexI][indexJ]
            Macno //= 10
            JE -= 1

        Cdigit = Mchkdigit % 10
        return Cdigit
# ...existing code...



import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.utils.exceptions import InvalidFileException
from datetime import datetime
import os

# Define the single Excel file and sheet names
DATA_FILE = "atm_data.xlsx"
BIN_SHEET_NAME = "bin_table"
DISPUTES_SHEET_NAME = "disputes"
# Constant for the new ATM ID validation sheet name
ATM_IDS_SHEET_NAME = "valid_atmid"


def populate_bin_sheet(filename, sheet_name, data):
    """
    Populates the specified sheet with BIN data.
    """
    try:
        wb = load_workbook(filename)
        sheet = wb[sheet_name]
        
        sheet.delete_rows(1, sheet.max_row)
        
        sheet.append(["BIN", "FIID", "BANK_NAME"])
        
        for bin_value, fiid_value, bank_name in data:
            sheet.append([bin_value, fiid_value, bank_name])
            
        wb.save(filename)
    except Exception as e:
        print("Failed to populate '{sheet_name}': {e}")

def create_initial_excel_file():
    """Creates the main data file and necessary sheets if they do not exist."""
    if not os.path.exists(DATA_FILE):
        wb = Workbook()
        if 'Sheet' in wb.sheetnames and len(wb['Sheet'].tables) == 0 and wb['Sheet'].cell(1, 1).value is None:
             wb.remove(wb['Sheet'])

        wb.create_sheet(BIN_SHEET_NAME)
        wb.create_sheet(DISPUTES_SHEET_NAME)
        # Create the new ATM IDs sheet with the required header
        atm_ids_sheet = wb.create_sheet(ATM_IDS_SHEET_NAME)
        atm_ids_sheet.append(["ATMID"]) 
        
        disputes_sheet = wb[DISPUTES_SHEET_NAME]
        headers = ["ref", "txndate", "cardno", "acno", "atmid", "txnno", "amount", "branch",
                   "debit_ac", "credit_ac", "status", "posting_date", "file_type",
                   "remarks", "posting_flag", "posting_user", "card_fiid", "term_fiid",
                   "comp_type", "disp_type", "src_ip", "user_name"]
        disputes_sheet.append(headers)
        
        wb.save(DATA_FILE)
        print(f"Created initial data file: {DATA_FILE}")

      


def load_sheet_dict(filename, sheet_name, key_column):
    """
    Loads data from a specific sheet into a dictionary/set for quick lookup.
    Returns a set for the valid_atmid sheet.
    """
    try:
        wb = load_workbook(filename)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return {}
        sheet = wb[sheet_name]
        
        if sheet_name == ATM_IDS_SHEET_NAME:
            data_set = set()
            # Iterate through rows starting from the second row (skipping header)
            for row in sheet.iter_rows(min_row=2, values_only=True):
                # Ensure the cell has a value before adding it
                if row and row is not None:
                    # Strip whitespace and convert to uppercase for robust matching
                    data_set.add(str(row[0]).strip().upper())
            return data_set
        
        # Original logic for multi-column sheets like bin_table
        data_dict = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            data_key = row[key_column]
            if data_key is not None:
                data_dict[str(data_key)] = row 
        wb.close()
        return data_dict
    except FileNotFoundError:
        return {}
    except InvalidFileException:
        return {}
    except IndexError:
        return {}


def append_fo_atm_record(record):
    """Appends a dispute record to the 'disputes' sheet."""
    try:
        wb = load_workbook(DATA_FILE)
        sheet = wb[DISPUTES_SHEET_NAME]
    except (FileNotFoundError, InvalidFileException):
        messagebox.showerror("Error", f"{DATA_FILE} not found or corrupted.")
        return

    sheet.append(record)
    wb.save(DATA_FILE)
    wb.close()

# -----------------------------------------------------------
# REPLACED CODE BLOCK: New ATM ID Validation Logic
# -----------------------------------------------------------

def is_valid_atm_id(atmid, valid_atm_ids_set):
    """
    Validates the ATM ID by checking its presence in the provided set of valid IDs 
    loaded from the 'valid_atmid' Excel sheet.
    """
    # Simply check if the user input exists in the set of valid IDs
    if atmid.upper() in valid_atm_ids_set:
        return True, "" # Valid ID found in the list, no error message needed
    else:
        return False, "ATM ID is invalid or not found in the approved list."

# -----------------------------------------------------------

# This function remains unchanged, as it expects a structured ATM ID format 
# to extract the 5th character's information.
def get_term_fiid(atmid):
    term_bank = atmid[4]
    if term_bank == "F":
        return atmid[4:8]
    elif term_bank in "0123457":
        mapping = {"0": "C001", "1": "C021", "2": "C022", "3": "C023", "4": "C024", "5": "C025", "7": "C027"}
        return mapping.get(term_bank, "NIL")
    else:
        return "NIL"

def get_card_fiid(cardno, bin_table):
    sbin = cardno[:6]
    fbin = cardno[:9]
    card_bank = cardno[6:7]
    if sbin == "622018":
        if card_bank in ["0", "6", "3", "1"]:
            return "C001"
    if fbin in bin_table:
        return bin_table[fbin][1]
    if sbin in bin_table:
        return bin_table[sbin][1]
    return None

def format_txn_date(txndate):
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            dt = datetime.strptime(txndate, fmt)
            return dt.strftime("%d-%b-%y").upper()
        except:
            pass
            print(f" Invalid date format entered: {txndate}. Using default date 31-DEC-99.")
        return "31-DEC-99"

def validate_acno(acno):
    while len(acno) > 11 and acno.startswith("0"):
        acno = acno[1:]
    return acno

class ATMDisputeApp:
    def __init__(self, root):
        self.root = root
        root.title("ATM Dispute Processing")
        
        create_initial_excel_file() 
        
        self.bin_table = load_sheet_dict(DATA_FILE, BIN_SHEET_NAME, 0)
        self.valid_atm_ids_set= load_sheet_dict(DATA_FILE, ATM_IDS_SHEET_NAME, 0)
        self.ac_details = load_sheet_dict(DATA_FILE, "acdetails", 0)
      

         # Define all fields and their types (Entry or Radio)
        fields_config = [
            ("Ref", "Entry"), 
            ("Txn Date (DD/MM/YYYY)", "Entry"), 
            ("Card No", "Entry"), 
            ("Ac No", "Entry"), 
            ("ATM ID", "Entry"), 
            ("Txn No", "Entry"),
            ("Amount", "Entry"), 
            ("Comp Type", "Radio", ["SOF", "FOS", "FOF"]), 
            ("Credit To", "Radio", ["cust", "other"]), 
            ("Disp Type", "Radio", ["dd", "short", "unsucc", "full"])
        ]

        for i, field_conf in enumerate(fields_config):
            self.messagebox =tk.Text(root, height=10, width=80)
            self.messagebox.grid(row=len(fields_config)+2, column=0, columnspan=3, pady=5, padx=5, sticky="nsew")
            self.messagebox.config(state=tk.DISABLED)

        
            

        if not self.valid_atm_ids_set:
            self.log_message("Warning: '{ATM_IDS_SHEET_NAME}' sheet is empty. ATM ID validation will fail.")
        
        # Data structure to hold input widgets
        self.entries = {}
        # Data structures to hold Radiobutton variables
        self.comp_type_var = tk.StringVar(value="SOF") # Default value
        self.credit_to_var = tk.StringVar(value="cust") # Default value
        self.disp_type_var = tk.StringVar(value="dd") # Default value

       
        # Layout the UI elements
        for i, field_conf in enumerate(fields_config):
            label_text = field_conf[0]
            label = tk.Label(root, text=label_text)
            label.grid(row=i, column=0, sticky='e', pady=2, padx=5)

            if field_conf[1] == "Entry":
                entry = tk.Entry(root, width=40)
                entry.grid(row=i, column=1, columnspan=2, sticky='w', pady=2, padx=5)
                self.entries[label_text] = entry
            elif field_conf[1] == "Radio":
                options = field_conf[2]
                # Determine which StringVar to use
                if label_text == "Comp Type":
                    var = self.comp_type_var
                elif label_text == "Credit To":
                    var = self.credit_to_var
                elif label_text == "Disp Type":
                    var = self.disp_type_var
                
                # Create a frame to hold the radio buttons horizontally
                frame = tk.Frame(root)
                frame.grid(row=i, column=1, columnspan=2, sticky='w', pady=2, padx=5)
                for j, option in enumerate(options):
                    rb = tk.Radiobutton(frame, text=option, variable=var, value=option)
                    rb.pack(side=tk.LEFT, padx=5)
                # Store the variable reference for easy access later
                self.entries[label_text] = var


        self.user_name = "DefaultUser" 
        tk.Button(root, text="Process Dispute", command=self.process_dispute).grid(row=len(fields_config)+1, column=1, pady=10, columnspan=2)

        self.messagebox = tk.Text(root, height=10, width=80)
        self.messagebox.grid(row=len(fields_config)+2, column=0, columnspan=3, pady=5, padx=5)
        self.messagebox.config(state=tk.DISABLED)
        
        if not self.bin_table:
            self.log_message(f"Warning: '{BIN_SHEET_NAME}' sheet in '{DATA_FILE}' is empty. BIN validation may fail.")

    def log_message(self, msg):
        self.messagebox.config(state=tk.NORMAL)
        self.messagebox.insert(tk.END, msg + "\n")
        self.messagebox.see(tk.END)
        self.messagebox.config(state=tk.DISABLED)

    def process_dispute(self):
        # Extract data from Entries and RadioButton variables
        ref = self.entries["Ref"].get() or "Not Known"
        txndate = self.entries["Txn Date (DD/MM/YYYY)"].get()
        cardno = self.entries["Card No"].get().strip()
        acno = self.entries["Ac No"].get().strip()
        atmid = self.entries["ATM ID"].get().strip().upper()
        txnno = self.entries["Txn No"].get().strip()
        amount = self.entries["Amount"].get().strip()
        
        # Get values from StringVars tied to Radiobuttons
        comp_type = self.comp_type_var.get()
        credit_to = self.credit_to_var.get()
        disp_type = self.disp_type_var.get()

        # Basic input validation
        if not cardno or not atmid or not amount:
            messagebox.showerror("Validation Error", "Card No, ATM ID, and Amount are required fields.")
            return

        acno = validate_acno(acno)
        if not txndate:
            txndate = "31/12/1999"
        txndate_fmt = format_txn_date(txndate)

        user_name = self.user_name 

        valid_atm, rej_msg = is_valid_atm_id(atmid, self.valid_atm_ids_set)
        if not valid_atm:
            messagebox.showerror("Validation Error", rej_msg)
            return

        term_fiid = get_term_fiid(atmid)
        card_fiid = get_card_fiid(cardno, self.bin_table)
        if card_fiid is None:
            messagebox.showerror("Validation Error", "Invalid BIN or BIN not found in table.")
            return

        if term_fiid == card_fiid:
            messagebox.showerror("Validation Error", "Card and Term FIIDs are same, which is not allowed for this dispute type.")
            return

        branch = cardno[6:11]
        branch = branch.zfill(5)

        debit1, credit1 = "", ""
        file_type1 = "T1"

        record = [
            ref, txndate_fmt, cardno, acno, atmid, txnno, amount, branch,
            debit1, credit1, "NO", None, file_type1,
            "File1 Entry", "NO", user_name, card_fiid, term_fiid,
            comp_type, disp_type, "127.0.0.1", user_name 
        ]

        append_fo_atm_record(record)

        self.log_message("Dispute processed successfully.")
        self.log_message(f"Record for ref {ref} appended to {DATA_FILE}")

if __name__ == "__main__":
    create_initial_excel_file() 
    root = tk.Tk()
    app = ATMDisputeApp(root)
    root.mainloop()
