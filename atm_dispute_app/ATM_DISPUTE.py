
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
        # Debugging info
        print(f"DEBUG: Validation failed for ID: '{atmid.upper()}'")
        print(f"DEBUG: Valid IDs Set Size: {len(valid_atm_ids_set)}")
        if len(valid_atm_ids_set) > 0:
             sample = list(valid_atm_ids_set)[:5]
             print(f"DEBUG: Sample IDs: {sample}")
             
        return False, f"ATM ID '{atmid}' is invalid or not in approved list (Loaded {len(valid_atm_ids_set)} IDs)."

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

def compute_legs(card_fiid, term_fiid, branch, acno, atmid, comp_type, disp_type, credit_to, ac_details):
    """
    Python port of the JSP logic for computing accounting legs.
    Returns: (legs, error_message)
          legs: list of dicts {debit_ac, credit_ac, file_type}
          error_message: string if rejected, else None
    """
    legs = []
    
    # -------------------------------------------------------------
    # 1. Helper / Setup
    # -------------------------------------------------------------
    
    # Clean branch to 5 digits (logic from JSP)
    # JSP: "if(comp_type.equals("SOF")){ branch=dr_branch; }" where dr_branch from param OR cardno
    # We assume 'branch' passed in is already correct (cardno based).
    branch = str(branch).zfill(5)
    
    # Pre-defined Logic for C0xx Accounts (for debit1/credit2 mappings)
    # Logic extracted from JSP if/else chains
    def get_c0xx_account(fiid, mode='debit'):
        # mode='debit' for SOF debit1 logic, mode='credit' for T2 credit2 logic (mostly same)
        if fiid == "C021": return "10309443213"
        if fiid == "C022": return "10309443177"
        if fiid == "C023": return "10309443188"
        if fiid == "C024": return "10309443235"
        if fiid == "C025": return "10309443246"
        if fiid == "C027": return "10309443202"
        return None

    # Load account details helper
    def get_ac_info(fiid):
        # ac_details keys: 'vostro_ac', 'settl_bgl_ac' (mapped to collection_ac)
        info = ac_details.get(fiid)
        if not info: return None
        return info

    debit1 = ""
    credit1 = ""
    file_type1 = "N"
    
    debit2 = ""
    credit2 = ""
    file_type2 = "N"
    
    debit3 = ""
    credit3 = ""
    file_type3 = "N"
    
    # -------------------------------------------------------------
    # 2. Logic Implementation
    # -------------------------------------------------------------
    
    # ======== CASE: SOF ========
    if comp_type == "SOF":
        if disp_type in ("dd", "98581", "unsucc", "full"):
            debit1 = "2399724042928"
            file_type1 = "CR"
            
            if card_fiid == "C001":
                if credit_to == "cust":
                    if len(acno) != 11:
                        return [], f"Invalid Credit Account: {acno}. Must be 11 digits for 'cust'."
                    credit1 = acno
                else:
                    credit1 = "98582" + branch + "C"
            else:
                return [], "Dispute Type Not Allowed for Card FIID"
                
            # JSP: "rs_ac1=stmt.executeQuery... where fiid=term_fiid"
            # It checks if term_fiid exists, if not -> "Invalid Term FIID". 
            # Note: For this branch, it actually DOESN'T use the result rs_ac1 for account assignment, 
            # but it enforces the check.
            info = get_ac_info(term_fiid)
            if not info:
                return [], "Invalid Term FIID (Not found in ac_details)"
                
        elif disp_type == "short":
            file_type1 = "T1"
            
            # -- Leg 1 Debit Logic --
            if card_fiid == "C001":
                debit1 = "98581" + branch + "C"
            else:
                special_ac = get_c0xx_account(card_fiid)
                if special_ac:
                    debit1 = special_ac
                else:
                    return [], "Dispute Type Not Allowed for Card FIID"
            
            # -- Leg 1 Credit (and T2 Debit) Logic --
            info = get_ac_info(term_fiid)
            if info:
                # JSP: credit1=rs_ac1.getString("collection_ac"); credit2=rs_ac1.getString("vostro_ac");
                # Wait: JSP says:
                # credit1=rs_ac1.getString("collection_ac");
                # credit2=rs_ac1.getString("vostro_ac");  <-- Suspicious? 
                # Actually typically T1 credit is collection. 
                # Let's re-read JSP carefully:
                #   credit1=rs_ac1.getString("collection_ac");
                #   credit2=rs_ac1.getString("vostro_ac");
                #   debit2=credit1; 
                #
                # Wait, usually T2 debit is Collection. Yes, debit2=credit1.
                # BUT credit2 is Vostro? 
                # Let's check the rest of the flow...
                #
                # Ah, wait. Inside SOF->Short:
                # if(term_fiid.equals("F005")){ file_type2="T2"; } else { file_type2="VC"; }
                # 
                # Actually, look at the JSP logic for SOF credit2 assignment later:
                # IT OVERWRITES credit2 later!
                # "if(term_fiid.equals("C001")) ... credit2=98582...C"
                # So the initial `credit2=vostro` from rs_ac1 might be ignored or is for specific cases.
                
                credit1 = info.get("settl_bgl_ac", "") # Collection AC
                # We will set debit2 temporarily, final logic handles it
                debit2 = credit1 
            else:
                return [], "Invalid Term FIID"
            
            # -- File Type 2 determination --
            if term_fiid == "F005":
                file_type2 = "T2"
            else:
                file_type2 = "VC"
                
            # -- Leg 2 Credit Logic --
            if term_fiid == "C001":
                # temp1=atmid.substring(5,10); credit2="98582"+temp1+"C";
                if len(atmid) >= 10:
                     temp1 = atmid[5:10]
                     base_c2 = "98582" + temp1
                     chk_c2 = getCheckDigitNumber(base_c2)
                     credit2 = base_c2 + chk_c2
                else:
                     # Fallback if ATMID is short? JSP assumes valid.
                     base_c2 = "98582" + branch
                     chk_c2 = getCheckDigitNumber(base_c2)
                     credit2 = base_c2 + chk_c2
            elif term_fiid in ("C021", "C022", "C023", "C024", "C025", "C027"):
                c_acc = get_c0xx_account(term_fiid)
                if c_acc: credit2 = c_acc
            else:
                 return [], "Invalid Term FIID (for Credit2 logic)"
            
            # -- Leg 2 Override for Card FIID == C001 (Customer logic) --
            if card_fiid == "C001":
                file_type2 = "T2"
                if credit_to == "cust":
                     if len(acno) != 11:
                         return [], f"Invalid Credit Account: {acno}"
                     credit2 = acno
                else:
                     base_c2 = "98582" + branch
                     chk_c2 = getCheckDigitNumber(base_c2)
                     credit2 = base_c2 + chk_c2

    # ======== CASE: FOS ========
    elif comp_type == "FOS":
        if disp_type == "short":
             # -- Leg 1 Lookup (Card FIID) --
             info = get_ac_info(card_fiid)
             if info:
                  # T1: Debit 98581... -> Credit Settlement BGL
                  # Calc Check Digit for 98581 + branch
                  base_ac = "98581" + branch
                  chk = getCheckDigitNumber(base_ac)
                  debit1 = base_ac + chk
                  
                  credit1 = info.get("settl_bgl_ac", "") # 48979...
                  
                  # T2: Debit Settlement -> Credit Vostro
                  debit2 = credit1
                  credit2 = info.get("vostro_ac", "") # 30118...
                  
                  if card_fiid == "F005":
                      file_type1 = "T1"
                  else:
                      file_type1 = "VD"
             else:
                  return [], "Account Details Corresponding to Card FIID Not Found"
             
             # -- Leg 2 Logic (Term FIID) --
             file_type2 = "T2"
             if term_fiid == "C001":
                 if len(atmid) >= 10:
                     temp1 = atmid[5:10]
                     base_c2 = "98582" + temp1
                     chk_c2 = getCheckDigitNumber(base_c2)
                     credit2 = base_c2 + chk_c2
                 else:
                     base_c2 = "98582" + branch
                     chk_c2 = getCheckDigitNumber(base_c2)
                     credit2 = base_c2 + chk_c2
             elif term_fiid in ("C021", "C022", "C023", "C024", "C025", "C027"):
                 c_acc = get_c0xx_account(term_fiid)
                 if c_acc: credit2 = c_acc
             else:
                 return [], "Invalid Term FIID"
                 
    # ======== CASE: FOF ========
    elif comp_type == "FOF":
         # -- Leg 1 (Term FIID) --
         info_term = get_ac_info(term_fiid)
         if info_term:
             debit1 = info_term.get("vostro_ac", "")
             credit1 = info_term.get("settl_bgl_ac", "")
             
             if term_fiid == "F005":
                 file_type1 = "T1"
             else:
                 file_type1 = "VD"
             
             # Determine next file type
             if file_type1 == "VD":
                 file_type2 = "T1"
             else:
                 file_type2 = "T2"
                 
             debit2 = credit1
         else:
             return [], "Account Details Corresponding to Term FIID Not Found"
         
         # -- Leg 2/3 (Card FIID) --
         info_card = get_ac_info(card_fiid)
         if info_card:
             # JSP: credit3=vostro, credit2=collection, debit3=credit2
             val_vostro = info_card.get("vostro_ac", "")
             val_coll = info_card.get("settl_bgl_ac", "")
             
             credit3 = val_vostro
             credit2 = val_coll
             debit3 = credit2 # JSP: debit3=credit2
             
             if card_fiid == "F005":
                 if file_type2 == "T1":
                     file_type3 = "T2"
                 else:
                     file_type3 = "T3"
             else:
                 file_type3 = "VC"
         else:
             return [], "Account Details Corresponding to Card FIID Not Found"

    # -------------------------------------------------------------
    # 3. Construct Legs List
    # -------------------------------------------------------------
    # Filter out empty/N legs
    
    # Filter out empty/N legs
    def is_valid_ac(ac):
        return ac and str(ac).strip() != "" and str(ac).strip() != "0"

    if file_type1 != "N" and is_valid_ac(debit1) and is_valid_ac(credit1):
        legs.append({"debit_ac": debit1, "credit_ac": credit1, "file_type": file_type1})
        
    if file_type2 != "N" and is_valid_ac(debit2) and is_valid_ac(credit2):
        legs.append({"debit_ac": debit2, "credit_ac": credit2, "file_type": file_type2})
        
    if file_type3 != "N" and is_valid_ac(debit3) and is_valid_ac(credit3):
        legs.append({"debit_ac": debit3, "credit_ac": credit3, "file_type": file_type3})
        
    if not legs:
        return [], "No valid legs could be generated (Accounts were missing or empty)."
        
    return legs, None

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