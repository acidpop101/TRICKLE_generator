
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
# Define the single Excel file and sheet names
DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "data", "atm_data.xlsx")
BIN_SHEET_NAME = "bin_table"
DISPUTES_SHEET_NAME = "disputes"
AC_DETAILS_SHEET_NAME = "ac_details"

def getCheckDigitNumber(acc_no):
    """Calculate check digit using 16x9 Matrix from constants.checkDigitConstantArray"""
    try:
        # Check if constant array is available (from imports)
        if 'checkDigitConstantArray' not in globals():
             print("Error: checkDigitConstantArray not imported.")
             return "0"

        Macno = int(acc_no)
        JE = 15
        Mchkdigit = 0
        
        while Macno > 0:
            iLastDigit = Macno % 10
            if iLastDigit > 0:
                Mdigit = iLastDigit - 1
                indexI = JE
                indexJ = Mdigit
                
                if 0 <= indexI < 16 and 0 <= indexJ < 9:
                    Mchkdigit += checkDigitConstantArray[indexI][indexJ]
            
            Macno //= 10
            JE -= 1
            
        return str(Mchkdigit % 10)
    except Exception as e:
        print(f"Error calculating check digit: {e}")
        return "0"


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
        wb.create_sheet(AC_DETAILS_SHEET_NAME) # Create ac_details sheet
        
        disputes_sheet = wb[DISPUTES_SHEET_NAME]
        headers = ["ref", "txndate", "cardno", "acno", "atmid", "txnno", "amount", "branch",
                   "debit_ac", "credit_ac", "status", "posting_date", "file_type"]
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
            
            if BIN_SHEET_NAME not in wb.sheetnames:
                wb.create_sheet(BIN_SHEET_NAME).append(["BIN", "FIID", "BANK_NAME"])
                save_needed = True

            if DISPUTES_SHEET_NAME not in wb.sheetnames:
                 disputes_sheet = wb.create_sheet(DISPUTES_SHEET_NAME)
                 headers = ["ref", "txndate", "cardno", "acno", "atmid", "txnno", "amount", "branch",
                            "debit_ac", "credit_ac", "status", "posting_date", "file_type"]
                 disputes_sheet.append(headers)
                 save_needed = True
            
            # Populate if empty
            if wb[AC_DETAILS_SHEET_NAME].max_row <= 1:
                populate_ac_details_sheet(wb)
                save_needed = True
                
            if save_needed:
                wb.save(DATA_FILE)
                print(f"Updated data file with missing sheets.")
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

# This function remains unchanged, as it expects a structured ATM ID format 
# to extract the 5th character's information.
def get_term_fiid(atmid):
    if len(atmid) < 5:
        return "NIL"
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
    if len(atmid) < 5: return "NIL", "Invalid ATM ID"
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

def compute_legs(card_fiid, term_fiid, branch, acno, atmid, comp_type, disp_type, credit_to, ac_details):
    legs = []
    
    branch = str(branch).zfill(5)
    atmid_temp1 = atmid[5:10] if len(atmid) >= 10 else branch
    
    def get_ac_info(fiid):
        return ac_details.get(fiid)

    file_type1, file_type2, file_type3 = "N", "N", "N"
    debit1, debit2, debit3 = "", "", ""
    credit1, credit2, credit3 = "", "", ""

    if term_fiid == "C001":
        file_type1 = "T1"
        debit1 = "98581" + atmid_temp1 + "C"
    elif term_fiid == "C021": file_type1, debit1 = "T1", "10309443213"
    elif term_fiid == "C022": file_type1, debit1 = "T1", "10309443177"
    elif term_fiid == "C023": file_type1, debit1 = "T1", "10309443188"
    elif term_fiid == "C024": file_type1, debit1 = "T1", "10309443235"
    elif term_fiid == "C025": file_type1, debit1 = "T1", "10309443246"
    elif term_fiid == "C027": file_type1, debit1 = "T1", "10309443202"

    if comp_type == "SOF":
        if term_fiid == "C001":
            debit1 = "2399724042928"
            file_type1 = "CR"
            if card_fiid == "C001":
                if credit_to == "cust":
                    if len(acno) != 11:
                        return [], f"Invalid Credit Account: {acno}"
                    credit1 = acno
                else:
                    credit1 = "98582" + branch + "C"
            else:
                return [], "Dispute Type Not Allowed for Card FIID"
        else:
            info = get_ac_info(term_fiid)
            if info:
                debit1 = info.get("vostro_ac", "")
                credit1 = info.get("settl_bgl_ac", "")
                debit2 = credit1
                if disp_type == "short":
                    file_type1 = "T1"
                    file_type2 = "VC"
                else:
                    file_type1 = "VD"
                    file_type2 = "VC"
            else:
                return [], "Account Details Corresponding to Term FIID Not Found"
                
            if card_fiid == "C001":
                file_type2 = "T2"
                if credit_to == "cust":
                    if len(acno) != 11:
                        return [], f"Invalid Credit Account: {acno}"
                    credit2 = acno
                else:
                    credit2 = "98582" + branch + "C"
            elif card_fiid == "C021": file_type2, credit2 = "T2", "10309443213"
            elif card_fiid == "C022": file_type2, credit2 = "T2", "10309443177"
            elif card_fiid == "C023": file_type2, credit2 = "T2", "10309443188"
            elif card_fiid == "C024": file_type2, credit2 = "T2", "10309443235"
            elif card_fiid == "C025": file_type2, credit2 = "T2", "10309443246"
            elif card_fiid == "C027": file_type2, credit2 = "T2", "10309443202"
            else:
                return [], "Invalid Card FIID"
                     
    elif comp_type == "FOS":
        file_1_type = "VD"
        file_2_type = "VC"

        if disp_type == "short":
             info = get_ac_info(card_fiid)
             if info:
                 credit1 = info.get("settl_bgl_ac", "")
                 debit2 = credit1
                 credit2 = info.get("vostro_ac", "")
                 file_type1 = "VD"
                 file_type2 = "VC"
             else:
                 return [], "Account Details Corresponding to Card FIID Not Found"
             
             if term_fiid == "C001": 
                 file_type1 = "T1"
                 file_type2 = "T2"
                 debit1 = "98581" + atmid_temp1 + "C"
             elif term_fiid == "C021": credit2 = "10309443213"
             elif term_fiid == "C022": credit2 = "10309443177"
             elif term_fiid == "C023": credit2 = "10309443188"
             elif term_fiid == "C024": credit2 = "10309443235"
             elif term_fiid == "C025": credit2 = "10309443246"
             elif term_fiid == "C027": credit2 = "10309443202"
             else: return [], "Invalid Term FIID"
             
        else:
             info = get_ac_info(card_fiid)
             if info:
                 credit2 = info.get("vostro_ac", "")
                 credit1 = info.get("settl_bgl_ac", "")
                 debit2 = credit1
                 file_type2 = file_2_type
             else:
                 return [], "Account Details Corresponding to Card FIID Not Found"
             
             if term_fiid == "C001":
                 file_type1 = "T1"
                 file_type2 = "T2"
                 debit1 = "98581" + atmid_temp1 + "C"
             elif term_fiid == "C021": file_type1, debit1 = file_1_type, "10309443213"
             elif term_fiid == "C022": file_type1, debit1 = file_1_type, "10309443177"
             elif term_fiid == "C023": file_type1, debit1 = file_1_type, "10309443188"
             elif term_fiid == "C024": file_type1, debit1 = file_1_type, "10309443235"
             elif term_fiid == "C025": file_type1, debit1 = file_1_type, "10309443246"
             elif term_fiid == "C027": file_type1, debit1 = file_1_type, "10309443202"
             else: return [], "Invalid Term FIID"
                 
    elif comp_type == "FOF":
         info_term = get_ac_info(term_fiid)
         if info_term:
             debit1 = info_term.get("vostro_ac", "")
             credit1 = info_term.get("settl_bgl_ac", "")
             file_type1 = "VD"
             file_type2 = "T1"
             debit2 = credit1
         else:
             return [], "Account Details Corresponding to Term FIID Not Found"
         
         info_card = get_ac_info(card_fiid)
         if info_card:
             credit3 = info_card.get("vostro_ac", "")
             credit2 = info_card.get("settl_bgl_ac", "")
             debit3 = credit2
             file_type3 = "VC"
         else:
             return [], "Account Details Corresponding to Card FIID Not Found"

         if term_fiid == "C001":
             file_type1 = "T1"
             debit1 = "98581" + atmid_temp1 + "C"
         elif term_fiid == "C021": file_type1, debit1 = "T1", "10309443213"
         elif term_fiid == "C022": file_type1, debit1 = "T1", "10309443177"
         elif term_fiid == "C023": file_type1, debit1 = "T1", "10309443188"
         elif term_fiid == "C024": file_type1, debit1 = "T1", "10309443235"
         elif term_fiid == "C025": file_type1, debit1 = "T1", "10309443246"
         elif term_fiid == "C027": file_type1, debit1 = "T1", "10309443202"


    def resolve_check_digit(ac):
        if str(ac).endswith("C"):
            base = str(ac)[:-1]
            return base + getCheckDigitNumber(base)
        return ac
        
    def is_valid_ac(ac):
        val = str(ac).strip()
        return val and val != "" and val != "0"

    debit1, credit1 = resolve_check_digit(debit1), resolve_check_digit(credit1)
    debit2, credit2 = resolve_check_digit(debit2), resolve_check_digit(credit2)
    debit3, credit3 = resolve_check_digit(debit3), resolve_check_digit(credit3)

    if file_type1 != "N" and is_valid_ac(debit1) and is_valid_ac(credit1):
        legs.append({"debit_ac": debit1, "credit_ac": credit1, "file_type": file_type1})
    if file_type2 != "N" and is_valid_ac(debit2) and is_valid_ac(credit2):
        legs.append({"debit_ac": debit2, "credit_ac": credit2, "file_type": file_type2})
    if file_type3 != "N" and is_valid_ac(debit3) and is_valid_ac(credit3):
        legs.append({"debit_ac": debit3, "credit_ac": credit3, "file_type": file_type3})
        
    if not legs:
        return [], "No valid legs could be generated."
        
    return legs, None


class ATMDisputeApp:
    def __init__(self, root):
        self.root = root
        root.title("ATM Dispute Processing")
        
        create_initial_excel_file() 
        
        self.bin_table = load_sheet_dict(DATA_FILE, BIN_SHEET_NAME, 0)
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

        if comp_type == "FOS":
             # Use ATM Branch (Indices 5-10 e.g. T1BW000177... -> 00177)
             if len(atmid) >= 10:
                 branch = atmid[5:10]
             else:
                 branch = atmid[-5:] # Fallback
        else:
             branch = cardno[6:11]
        
        branch = branch.zfill(5)

        # NEW: compute legs (like JSP)
        legs, error_msg = compute_legs(
            card_fiid=card_fiid,
            term_fiid=term_fiid,
            branch=branch,
            acno=acno,
            atmid=atmid,
            comp_type=comp_type,
            disp_type=disp_type,
            credit_to=credit_to,
            ac_details=self.ac_details,
        )

        if error_msg:
             messagebox.showerror("Processing Error", error_msg)
             return

        # For EACH leg, write one row into disputes sheet
        for idx, leg in enumerate(legs, start=1) :
            debit_ac = leg["debit_ac"]
            credit_ac = leg["credit_ac"]
            file_type = leg["file_type"]
            
            record = [
                ref, txndate_fmt, cardno, acno, atmid, txnno, amount, branch,
                debit_ac, credit_ac, "NO", None, file_type
            ]

            append_fo_atm_record(record)

        self.log_message("Dispute processed successfully.")
        self.log_message(f"{len(legs)} leg(s) appended to {DATA_FILE} for ref {ref}")

if __name__ == "__main__":
    create_initial_excel_file() 
    root = tk.Tk()
    app = ATMDisputeApp(root)
    root.mainloop()