
import constants
import os
import sys
import openpyxl
from pathlib import Path
from openpyxl.utils.exceptions import InvalidFileException

# Import core classes
from file_object import FileObject
from record_object import RecordObject
from input_file_validation_rules import InputFileValidationRules
from report import Report
import error_code as ec

DISPUTES_SHEET_NAME = "disputes"
AC_DETAILS_SHEET_NAME = "ac_details"

def load_known_bgl_accounts(wb):
    """
    Loads all BGL/Vostro accounts from ac_details sheet and returns a set.
    Also includes known static BGLs.
    """
    bgl_set = {
        "2399724042928", # BGL_POS_PAYABLE
    }
    
    if AC_DETAILS_SHEET_NAME in wb.sheetnames:
        sheet = wb[AC_DETAILS_SHEET_NAME]
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None: continue
            # row is (FIID, VOSTRO_AC, SETTL_BGL_AC)
            vostro = str(row[1] or "").strip()
            settl = str(row[2] or "").strip()
            if vostro: bgl_set.add(vostro)
            if settl: bgl_set.add(settl)
            
    return bgl_set

def leg_to_cbs_account_and_type(file_type, debit_ac, credit_ac, comp_type, disp_type, bgl_set=None):
    """
    Decide which account goes into the CBS record and what CBS account type to use (01/04/12/51/54/62).
    """
    if bgl_set is None: bgl_set = set()

    # For now:
    # - Use debit_ac as the posting account for DR legs
    # - Use credit_ac as the posting account for CR legs
    # - Map by file_type: T1/T2/VD/VC/CR
    
    # Default values
    is_customer_leg = True
    is_credit = False
    acct = debit_ac or credit_ac or ""
    acct_str = str(acct).strip()

    # Simple heuristic:
    if file_type == "CR":
        is_credit = True
        acct = credit_ac or debit_ac
    elif file_type in ("T1", "VD"):
        # Treat as debit legs (BGL or customer)
        is_credit = False
        acct = debit_ac or credit_ac
    elif file_type in ("T2", "VC"):
        # Treat as credit legs
        is_credit = True
        acct = credit_ac or debit_ac

    # DECIDE Account Type logic:
    # Check against known BGL set OR standard patterns for BGLs (98581... / 98582...)
    # Handle leading zeros which might be present in acct_str
    clean_acct = acct_str.lstrip('0')
    # Hardcoded BGLs from ATM_DISPUTE logic
    known_bgl_list = [
        "10309443213", "10309443177", "10309443188", 
        "10309443235", "10309443246", "10309443202",
        "2399724042928"
    ]
    
    is_bgl = (acct_str in bgl_set) or \
             (clean_acct in bgl_set) or \
             (clean_acct in known_bgl_list) or \
             clean_acct.startswith("9858") or \
             clean_acct.startswith("10309") or \
             (not clean_acct.isdigit()) # Ultimate safety: if it has letters, it MUST be BGL
             
    # If it is a BGL, it is NOT a customer leg
    is_customer_leg = not is_bgl

    # Map to CBS type
    if not is_customer_leg and not is_credit:
        cbs_type = "54"     # BGL_DR
    elif not is_customer_leg and is_credit:
        cbs_type = "04"     # BGL_CR
    elif is_customer_leg and not is_credit:
        cbs_type = "51"     # CUSTOMER_DR
    else:
        cbs_type = "01"     # CUSTOMER_CR

    return acct, cbs_type

def process_file(args):
    inputDirectoryPath = args[0] if len(args) > 0 else ""
    outputDirectoryPath = args[1] if len(args) > 1 else ""
    outputReportName = args[2] if len(args) > 2 else "REPORT.txt"

    report = Report()
    report.setReportName(outputReportName)
    report.setOutputDirectoryPath(outputDirectoryPath)
    report.setInputDirectoryPath(inputDirectoryPath)

    if not os.path.isdir(inputDirectoryPath):
        report.setError(True)
        report.setErrorCode("INPUT DIRECTORY DOES NOT EXIST.")
        print("INPUT DIRECTORY DOES NOT EXIST.")
        return report

    file_list = [
        os.path.join(inputDirectoryPath, f)
        for f in os.listdir(inputDirectoryPath)
        if os.path.isfile(os.path.join(inputDirectoryPath, f)) and f.lower().endswith('.dat')
    ]

    inputFileValidationRules = InputFileValidationRules()

    # ---------- START of loop over files ----------
    for file_path in file_list:
        if os.path.basename(file_path) == outputReportName:
            continue
            
        file_obj = FileObject()
        record = RecordObject()
        file_obj.setFileName(os.path.basename(file_path))

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                for line in f:
                    sCurrentLine = line.rstrip("\n\r")
                    record.setRecordToInitial()
                    record.setRecord(sCurrentLine)

                    inputFileValidationRules.inputFileRules(record, file_obj)
                    if record.isErrorFlag():
                        break
                    
                    file_obj.setRecordCount(file_obj.getRecordCount() + 1)
        except Exception as ex:
            print(f"File read error {file_path}: {ex}")

        # these two lines MUST be inside the for-loop
        inputFileValidationRules.setFileObjectWithRecordDetails(record, file_obj)
        report.getFileEntries().append(file_obj)
    # ---------- END of loop over files ----------

    report.setReportContent(report)

    out_dir = report.getOutputDirectoryPath()
    if out_dir:
        Path(out_dir).mkdir(parents=True, exist_ok=True)
        output_file_path = os.path.join(out_dir, report.getReportName())
        try:
            with open(output_file_path, "w", encoding="utf-8") as fw:
                fw.write(report.getReport())
        except Exception as ex:
            print("----------FILE NOT GENERATED SUCCESSFULLY----------")
            report.setError(True)
            report.setErrorCode(ec.FILE_WRITE_ERR)
            print(ex)

    return report
def excel_to_cbs_files(excel_path, output_dir):
    """Convert atm_data.xlsx → CBS fixed-width .dat files (CORRECTED FORMAT)"""
    try:
        wb = openpyxl.load_workbook(excel_path)
        if DISPUTES_SHEET_NAME not in wb.sheetnames:
            print(f"ERROR: No '{DISPUTES_SHEET_NAME}' sheet found")
            return None

        disputes_sheet = wb[DISPUTES_SHEET_NAME]
        
        # Load known BGL accounts for type deduction
        bgl_set = load_known_bgl_accounts(wb)
        
        files_data = {}  # Group by file_type

        for row_idx, row in enumerate(disputes_sheet.iter_rows(min_row=2, values_only=True), 2):
            if len(row) < 14 or row[12] is None:  # Skip incomplete rows
                continue
            
            # ... (rest of loop) ...


            ref = row[0]
            txndate = row[1]
            cardno = row[2]
            acno = row[3]
            atmid = row[4]
            txnno = row[5]
            amount_str = row[6]
            branch = row[7]
            debit_ac = row[8]
            credit_ac = row[9]
            status = row[10]
            posting_date = row[11]
            file_type = row[12]
            remarks = row[13]

            # Additional columns might exist
            posting_flag = row[14] if len(row) > 14 else None
            posting_user = row[15] if len(row) > 15 else None
            card_fiid = row[16] if len(row) > 16 else None
            card_bank_name = row[17] if len(row) > 17 else None
            term_fiid = row[18] if len(row) > 18 else None
            term_bank_name = row[19] if len(row) > 19 else None
            comp_type = row[20] if len(row) > 20 else None
            disp_type = row[21] if len(row) > 21 else None
            src_ip = row[22] if len(row) > 22 else None
            user_name = row[23] if len(row) > 23 else None

            if not file_type:
                continue

            # ✅ FIXED: Convert to paise with EXACT 16 digits
            try:
                amount_paise = int(float(amount_str or 0) * 100)
            except:
                amount_paise = 0
            amount_str_fixed = f"{amount_paise:016d}" # EXACTLY 16 digits, zero-padded
            
            leg_acct, account_type = leg_to_cbs_account_and_type(
                file_type=file_type,
                debit_ac=debit_ac,
                credit_ac=credit_ac,
                comp_type=comp_type,
                disp_type=disp_type,
                bgl_set=bgl_set
            )
            
            acno_clean = validate_acno_for_cbs(leg_acct)
            
            # Format Date YYMMDD
            pdate = "000101"
            if posting_date:
                pdate = posting_date or "000101"

            field_pad = " " * 10
            field_card = (str(cardno) if cardno else "").ljust(16)[:16]
            field_date = pdate[:6]
            field_txn  = (str(txnno) if txnno else "").zfill(9)[:9]
            field_br   = (str(branch) if branch else "").zfill(4)[:4]
            
            txn_details = (
                f"{field_pad}{field_card} {field_date} {field_txn} {field_br}"
            ).ljust(66)

            # ✅ CORRECT CBS Format
            cbs_record = (
                f"{account_type:<2}"
                f"{acno_clean:<17}"
                f"{amount_str_fixed:<16}"
                f"{txn_details:<66}"
            )
            
            if file_type not in files_data:
                files_data[file_type] = []
            files_data[file_type].append(cbs_record)






        wb.close()

        # Write .dat files per file_type
        cbs_files = {}
        for file_type, records in files_data.items():
            output_file = os.path.join(output_dir, f"{file_type.upper()}.dat")
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(records) + '\n')  # Add final newline
            cbs_files[file_type] = output_file
        
        return cbs_files

    except Exception as e:
        print(f"Excel to CBS conversion failed: {e}")
        import traceback
        traceback.print_exc()
        return None

def map_dispute_to_account_type(row):
    """Map dispute types to CBS account codes"""
    comp_type = row[20] if len(row) > 20 else "SOF"
    disp_type = row[21] if len(row) > 21 else "dd"

    mapping = {
        ("SOF", "dd"): "54",    # BGL_DR
        ("FOS", "short"): "62", # SYS_DR
        ("FOF", "unsucc"): "51",# CUSTOMER_DR
        ("SOF", "full"): "04",  # BGL_CR
        ("FOS", "dd"): "12",    # SYS_CR
    }
    return mapping.get((comp_type, disp_type), "01")  # Default CUSTOMER_CR

def validate_acno_for_cbs(acno):
    """FIXED: RIGHT-align with zeros for CBS format"""
    acno = str(acno or "").lstrip('0')[:13] # Clean to max 13 digits
    return acno.rjust(17, '0') # RIGHT align, zero pad to 17 chars

def validate_cbs_files(input_dir):
    """Validate generated CBS files (like original Trickle Feed)"""
    args = [input_dir, input_dir]  # Input=Output dir for validation
    report = process_file(args)
    print(report.getReport())
    return report

def main():
    if len(sys.argv) >= 3:
        input_path = sys.argv[1]
        output_path = sys.argv[2]
        
        # Check if input is Excel (ATM Dispute output)
        if input_path.lower().endswith('.xlsx'):
            print("Processing Excel -> CBS files...")
            cbs_files = excel_to_cbs_files(input_path, output_path)
            if cbs_files:
                print("CBS files generated:")
                for file_type, file_path in cbs_files.items():
                    print(f"  {file_path}")
                
                # Run validation on generated files
                validate_cbs_files(output_path)
            else:
                print("Excel processing failed")
        else:
            # Original folder processing
            print("Processing folder files...")
            process_file(sys.argv[1:])
    else:
        from file_process_gui import main as gui_main
        gui_main()

if __name__ == "__main__":
    main()
