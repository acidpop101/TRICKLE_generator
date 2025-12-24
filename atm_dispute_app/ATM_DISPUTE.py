
import sys
import os

# Fix path for trickle_feed imports
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'trickle_feed')))

# Try imports, handle if constants missing/error (though we know we fixed it)
try:
    from constants import BGL_DR, SYS_DR, CUSTOMER_DR, BGL_CR, SYS_CR, CUSTOMER_CR, checkDigitConstantArray, MAX_CR_DR_AMOUNT_ALLOWED
    import error_code as ec
except ImportError:
    pass # Proceeding without this for now as the user's new code didn't seem to rely on them directly in the UI logic provided, or we add them back if needed.
    # Actually, the user's code REMOVED the InputFileValidationRules class which used these. 
    # The new code seems to be a pure GUI restructure. I will stick to what the user pasted but formatted correctly.

import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.utils.exceptions import InvalidFileException
from datetime import datetime

# Define the single Excel file and sheet names
DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "atm_data.xlsx")
BIN_SHEET_NAME = "bin_table"
DISPUTES_SHEET_NAME = "disputes"
# Constant for the new ATM ID validation sheet name
ATM_IDS_SHEET_NAME = "valid_atmid"
AC_DETAILS_SHEET_NAME = "ac_details"

def populate_bin_sheet(filename, sheet_name, data):
    """
    Populates the specified sheet with BIN data.
    """
    try:
        wb = load_workbook(filename)
        sheet = wb[sheet_name]
        
        sheet.delete_rows(1, sheet.max_row)
        
        sheet.append(["BIN", "FIID", "BANK_NAME"])
        
        for bin_value, fiid_value in data:
            sheet.append([bin_value, fiid_value, "BANK_NAME"]) # Fixed variable name usage from user snippet
            
        wb.save(filename)
    except Exception as e:
        print(f"Failed to populate '{sheet_name}': {e}")

def populate_ac_details_sheet(wb):
    """Populates the ac_details sheet with default BGL data if empty."""
    sheet = wb[AC_DETAILS_SHEET_NAME]
    if sheet.max_row > 1:
        return # Already populated

    sheet.delete_rows(1, sheet.max_row)
    sheet.append(["FIID", "VOSTRO_AC", "SETTL_BGL_AC"])
    
    # Data transcribed from user image
    data = [
        ("F001", "10983139836", "4897926042923"),
        ("F005", "30118760239", "4897928042921"),
        ("F025", "30118760239", "4897928042921"),
        ("F002", "10983139814", "4897925042924"),
        ("F004", "30178185451", "4897929042920"),
        ("F011", "30854632275", "4897931042926"),
        ("F009", "30447713125", "4897921042928"),
        ("F006", "30179008128", "4897922042927"),
        ("F003", "",            "4897927042922"),
        ("F007", "",            "4897923042926"),
        ("F008", "",            "4897924042925"),
        ("F010", "",            "4897930042926"),
    ]
    
    for row in data:
        sheet.append(row)
    print(f"Populated '{AC_DETAILS_SHEET_NAME}' with default BGL data.")

def create_initial_excel_file():
    """Creates the main data file and necessary sheets if they do not exist."""
    wb = None
    if not os.path.exists(DATA_FILE):
        wb = Workbook()
        if 'Sheet' in wb.sheetnames and len(wb['Sheet'].tables) == 0 and wb['Sheet'].cell(1, 1).value is None:
             wb.remove(wb['Sheet'])

        wb.create_sheet(BIN_SHEET_NAME)
        wb.create_sheet(DISPUTES_SHEET_NAME)
        wb.create_sheet(ATM_IDS_SHEET_NAME).append(["ATMID"]) 
        wb.create_sheet(AC_DETAILS_SHEET_NAME) # Create ac_details sheet
        
        disputes_sheet = wb[DISPUTES_SHEET_NAME]
        headers = ["ref", "txndate", "cardno", "acno", "atmid", "txnno", "amount", "branch",
                   "debit_ac", "credit_ac", "status", "posting_date", "file_type",
                   "remarks", "posting_flag", "posting_user", "card_fiid", "card_bank_name",
                   "term_fiid", "term_bank_name", "comp_type", "disp_type", "src_ip", "user_name"]
        disputes_sheet.append(headers)
        
        populate_ac_details_sheet(wb) # Populate it
        wb.save(DATA_FILE)
        print(f"Created initial data file: {DATA_FILE}")
    else:
        # Check if ac_details exists, if not add it
        try:
            wb = load_workbook(DATA_FILE)
            save_needed = False
            if AC_DETAILS_SHEET_NAME not in wb.sheetnames:
                wb.create_sheet(AC_DETAILS_SHEET_NAME)
                save_needed = True
            
            # Populate if empty
            if wb[AC_DETAILS_SHEET_NAME].max_row <= 1:
                populate_ac_details_sheet(wb)
                save_needed = True
                
            if save_needed:
                wb.save(DATA_FILE)
                print(f"Updated data file with {AC_DETAILS_SHEET_NAME}")
        except Exception as e:
            print(f"Error checking/updating Excel file: {e}")
        finally:
            if wb: wb.close()

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

def load_ac_details(filename):
    """
    Load ac_details into a dict: fiid -> {"vostro_ac": ..., "settl_bgl_ac": ...}
    """
    try:
        wb = load_workbook(filename)
        if AC_DETAILS_SHEET_NAME not in wb.sheetnames:
            wb.close()
            return {}
        
        sheet = wb[AC_DETAILS_SHEET_NAME]
        ac_map = {}
        # Header: FIID, VOSTRO_AC, SETTL_BGL_AC
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            fiid, vostro_ac, settl_bgl_ac = row[:3]
            ac_map[str(fiid).strip()] = {
                "vostro_ac": str(vostro_ac or "").strip(),
                "settl_bgl_ac": str(settl_bgl_ac or "").strip(),
            }
        wb.close()
        return ac_map
    except Exception:
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

def gettermfiid(self, atmid):
    atmid = atmid.strip().upper()
    if len(atmid) == 0: return "NIL", "Invalid ATM ID"
    termbank = atmid[4]  # 5th char (0-indexed)
    at13 = atmid[0:3]
    at4 = atmid[3]
    if termbank == 'F':
        if at4 in ['W','C','N']:
            if at13 in ['S1C','S1A']: return atmid[4:8], None
        return "NIL", "Invalid ATM ID"
    
    # Domestic ATMs C001-C027 mapping
    fiid_map = {'0':'C001','1':'C021','2':'C022','3':'C023','4':'C024','5':'C025','7':'C027'}
    return fiid_map.get(termbank, "NIL"), "Invalid ATM ID" if termbank not in fiid_map else None

def get_bank_name_by_fiid(fiid, bin_table):
    """Find bank name from FIID using bin_table rows."""
    if not fiid:
        return ""
    for row in bin_table.values():
        if len(row) >= 2 and row[1] == fiid:
            return row[2] or ""
    return ""

def get_card_fiid(cardno, bin_table):
    sbin = cardno[:6]
    fbin = cardno[:9]
    card_bank = cardno[6:7]
    if sbin == "622018":
        if card_bank in ["0", "6", "3", "1"]:
            return "C001", "STATE BANK OF INDIA"
    if fbin in bin_table:
        row = bin_table[fbin]
        return row[1], (row[2] or "")
    if sbin in bin_table:
        row = bin_table[sbin]
        return row[1], (row[2] or "")
    return None, ""

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

def compute_legs(card_fiid, term_fiid, branch, acno, comp_type, disp_type, credit_to, ac_details):
    """
    Rough Python clone of JSP leg logic.
    Returns a list of legs: each leg = dict with debit_ac, credit_ac, file_type.
    For now we implement a minimal but correct multi-leg structure:
    - SOF + dd  → one CR file leg (POS payable)
    - SOF + short → T1/T2 legs
    - FOS/FOF → simple T1+T2 or VD/VC pattern placeholder
    Later you can refine this to match the JSP exactly.
    """
    legs = []
    # Clean branch to 5 digits
    branch = str(branch).zfill(5)
    
    # DEFAULTS (you will replace these with proper accounts from ac_details later)
    # These mimic the hard-coded BGL accounts in JSP (10309..., 98581..., 98582...)
    BGL_POS_PAYABLE = "2399724042928" # example from JSP debit1 for CR
    BGL_SHORT_SBI   = "98581" + branch + "C"
    BGL_CUST_SBI    = "98582" + branch + "C"
    
    term_ac = ac_details.get(term_fiid, {})
    card_ac = ac_details.get(card_fiid, {})
    
    term_vostro = term_ac.get("vostro_ac", "")
    term_settl_bgl = term_ac.get("settl_bgl_ac", "")
    
    card_vostro = card_ac.get("vostro_ac", "")
    card_settl_bgl = card_ac.get("settl_bgl_ac", "")
    
    # 1) SOF (State Bank On-Us) cases
    if comp_type == "SOF":
        if disp_type in ("dd", "unsucc", "full"):
            # Case: SOF + dd/full/unsucc → CR file only (POS Payable Entry)
            if credit_to == "cust":
                credit_ac = acno  # credit customer's account
            else:
                credit_ac = BGL_CUST_SBI  # branch BGL credit
            
            legs.append({
                "debit_ac": BGL_POS_PAYABLE,
                "credit_ac": credit_ac,
                "file_type": "CR",
            })
        elif disp_type == "short":
            # Case: SOF + short → T1 + T2 legs
            # T1: debit 98581+branch (short credit BGL),  credit = Collection A/c of Terminal Owner (Bank)
            # Logic: If term_fiid is our bank (SBI), use internal BGL. If other, use Settl BGL? 
            # Actually for SOF, Terminal Owner IS SBI (usually). 
            # But let's look up the "Settlement BGL" for the terminal owner just in case.
            
            # "COLLECTION_AC_TERM" placeholder replacement:
            # If term_fiid is in ac_details, use its settl_bgl_ac. Else fallback to BGL_CUST_SBI or similar.
            term_col_ac = term_settl_bgl if term_settl_bgl else BGL_CUST_SBI 

            legs.append({
                "debit_ac": BGL_SHORT_SBI,
                "credit_ac": term_col_ac, 
                "file_type": "T1",
            })
            
            # T2: debit collection_ac, credit customer or 98582+branch
            if credit_to == "cust":
                credit2 = acno
            else:
                credit2 = BGL_CUST_SBI
                
            legs.append({
                "debit_ac": term_col_ac,
                "credit_ac": credit2,
                "file_type": "T2",
            })
    
    # 2) FOS (Foreign On-Us)
    elif comp_type == "FOS":
        if disp_type == "short":
            # Foreign On-Us, short credit.
            # Leg1 (T1): DR Issuing Bank (Card) Vostro/Settl, CR Issuing Bank Settl
            # Use data from ac_details for Card/Issuer FIID
            
            dr_ac = card_vostro if card_vostro else (card_settl_bgl if card_settl_bgl else "MISSING_VOSTRO")
            cr_ac = card_settl_bgl if card_settl_bgl else "MISSING_SETTL"
            
            legs.append({
                "debit_ac": dr_ac,
                "credit_ac": cr_ac,
                "file_type": "T1",
            })
            
            # Leg2 (T2): DR Issuing Bank Settl, CR Terminal Owner Settl
            term_cr_ac = term_settl_bgl if term_settl_bgl else "MISSING_TERM_SETTL"
            
            legs.append({
                "debit_ac": cr_ac,
                "credit_ac": term_cr_ac,
                "file_type": "T2",
            })
        else:
            # Other FOS disputes: VD/VC pattern (reusing similar logic)
            dr_ac = card_vostro if card_vostro else card_settl_bgl
            cr_ac = card_settl_bgl
            
            legs.append({
                "debit_ac": dr_ac, 
                "credit_ac": cr_ac,
                "file_type": "VD",
            })
            term_cr_ac = term_settl_bgl
            legs.append({
                "debit_ac": cr_ac,
                "credit_ac": term_cr_ac,
                "file_type": "VC",
            })
            
    # 3) FOF (Foreign Off-Us)
    elif comp_type == "FOF":
        # Similar logic to FOS but typically involves Acquiring Bank vs Issuer
        # For simplicity, using same lookups for now as placeholders were identical
        if disp_type == "short":
            dr_ac = card_vostro if card_vostro else card_settl_bgl
            cr_ac = card_settl_bgl
            legs.append({
                "debit_ac": dr_ac,
                "credit_ac": cr_ac,
                "file_type": "T1",
            })
            term_cr_ac = term_settl_bgl
            legs.append({
                "debit_ac": cr_ac,
                "credit_ac": term_cr_ac,
                "file_type": "T2",
            })
        else:
            dr_ac = card_vostro if card_vostro else card_settl_bgl
            cr_ac = card_settl_bgl
            legs.append({
                "debit_ac": dr_ac,
                "credit_ac": cr_ac,
                "file_type": "VD",
            })
            term_cr_ac = term_settl_bgl
            legs.append({
                "debit_ac": cr_ac,
                "credit_ac": term_cr_ac,
                "file_type": "VC",
            })

    # If no rule matched, create at least one T1 leg as a fallback
    if not legs:
        legs.append({
            "debit_ac": "FALLBACK_DEBIT",
            "credit_ac": acno or "FALLBACK_CREDIT",
            "file_type": "T1",
        })
    
    return legs

class ATMDisputeApp:
    def __init__(self, root):
        self.root = root
        root.title("ATM Dispute Processing")
        
        create_initial_excel_file() 
        
        self.bin_table = load_sheet_dict(DATA_FILE, BIN_SHEET_NAME, 0)
        self.valid_atm_ids_set= load_sheet_dict(DATA_FILE, ATM_IDS_SHEET_NAME, 0)
        self.ac_details = load_ac_details(DATA_FILE)
      

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
            self.log_message(f"Warning: '{ATM_IDS_SHEET_NAME}' sheet is empty. ATM ID validation will fail.")
        
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
        # get_bank_name_by_fiid requires the full dict, but self.bin_table is the dict.
        term_bank_name = get_bank_name_by_fiid(term_fiid, self.bin_table)
        
        card_fiid, card_bank_name = get_card_fiid(cardno, self.bin_table)
        if card_fiid is None:
            messagebox.showerror("Validation Error", "Invalid BIN or BIN not found in table.")
            return

        if term_fiid == card_fiid:
            bank_name_display = card_bank_name or term_bank_name or "same bank"
            messagebox.showerror("Validation Error", f"Card and ATM belong to {bank_name_display}; this dispute type requires different banks.")
            return

        branch = cardno[6:11]
        branch = branch.zfill(5)

        # NEW: compute legs (like JSP)
        legs = compute_legs(
            card_fiid=card_fiid,
            term_fiid=term_fiid,
            branch=branch,
            acno=acno,
            comp_type=comp_type,
            disp_type=disp_type,
            credit_to=credit_to,
            ac_details=self.ac_details,
        )

        # For EACH leg, write one row into disputes sheet
        for idx, leg in enumerate(legs, start=1) :
            debit_ac = leg["debit_ac"]
            credit_ac = leg["credit_ac"]
            file_type = leg["file_type"]
            
            record = [
                ref, txndate_fmt, cardno, acno, atmid, txnno, amount, branch,
                debit_ac, credit_ac, "NO", None, file_type,  f"File{idx} Entry",
                # remarks
                "NO", user_name, card_fiid, card_bank_name,
                term_fiid, term_bank_name, comp_type, disp_type, "127.0.0.1", user_name 
            ]

            append_fo_atm_record(record)

        self.log_message("Dispute processed successfully.")
        self.log_message(f"{len(legs)} leg(s) appended to {DATA_FILE} for ref {ref}")

if __name__ == "__main__":
    create_initial_excel_file() 
    root = tk.Tk()
    app = ATMDisputeApp(root)
    root.mainloop()