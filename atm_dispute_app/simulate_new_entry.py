
import openpyxl

PATH = r"d:\TRICKLE_generator\atm_dispute_app\atm_data.xlsx"
SHEET = "disputes"

def simulate():
    try:
        wb = openpyxl.load_workbook(PATH)
        ws = wb[SHEET]
        
        # New Record Format: Indices 0-12 Only (13 items)
        # Ref, TxnDate, Card, Ac, ATM, Txn, Amt, Br, Dr, Cr, Stat, PDate, Type
        new_row = [
            "DUMMY_REF", 
            "31-Dec-25", 
            "1234567812345678", 
            "0", 
            "TEST0001", 
            "9999", 
            "5000", 
            "00001", 
            "99999999999999999", # Debit 
            "88888888888888888", # Credit
            "NO", 
            None, # Posting Date (Blank by code)
            "T1"  # File Type (Index 12)
        ]
        
        # Verify length
        print(f"Appending row with {len(new_row)} columns: {new_row}")
        
        ws.append(new_row)
        wb.save(PATH)
        
        print("Scrubbed remaining data in row... (Checking actual file state)")
        
        # PROOF: Read it back
        wb = openpyxl.load_workbook(PATH)
        ws = wb[SHEET]
        last_row_vals = []
        for cell in ws[ws.max_row]:
            last_row_vals.append(cell.value)
            
        print(f"Saved Row Length: {len(last_row_vals)}")
        print(f"Saved Row Values: {last_row_vals}")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    simulate()
