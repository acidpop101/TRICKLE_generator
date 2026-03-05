
import constants
import os
import sys
import openpyxl
from datetime import datetime
from pathlib import Path
from openpyxl.utils.exceptions import InvalidFileException

# Import core classes
from file_object import FileObject
from record_object import RecordObject
from input_file_validation_rules import InputFileValidationRules
from report import Report
import error_code as ec

def excel_to_cbs_files(excel_path, output_dir):
    """Convert atm_data.xlsx -> CBS fixed-width .dat files"""
    try:
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb["disputes"]
        
        # Build Set of known BGLs from ac_details
        known_bgls = set()
        if "ac_details" in wb.sheetnames:
            for r in wb["ac_details"].iter_rows(min_row=2, values_only=True):
                if r and r[0]:
                    if r[1]: known_bgls.add(str(r[1]).strip())
                    if len(r) > 2 and r[2]: known_bgls.add(str(r[2]).strip())
                    
        def is_bgl(ac):
            ac_s = str(ac).strip()
            if ac_s in known_bgls: return True
            if ac_s.startswith(('98581', '98582', '10309')): return True
            if ac_s == '2399724042928': return True
            if len(ac_s) > 11: return True
            return False
            
        files_content = {} # filename -> list of lines

        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]: continue 
            
            # Indexes based on ATM_DISPUTE:
            # Ref(0), TxnDate(1), CardNo(2), AcNo(3), ATMID(4), TxnNo(5), Amount(6), Branch(7), 
            # DrAc(8), CrAc(9), Status(10), PostDate(11), FileType(12)
            
            # Safe str conversion
            def safe_str(val): return str(val).strip() if val is not None else ""
            
            dr_ac = safe_str(row[8])
            cr_ac = safe_str(row[9])
            posting_date = row[11] if len(row) > 11 else None
            
            try:
                raw_amt = float(row[6]) if row[6] is not None else 0.0
                amount_paise = int(raw_amt * 100)
            except:
                amount_paise = 0
                
            card = safe_str(row[2])
            
            raw_date = row[1]
            if isinstance(raw_date, datetime):
                date_str = raw_date.strftime("%d%m%y") 
            else:
                try:
                    dt = datetime.strptime(safe_str(raw_date), "%d-%b-%y")
                    date_str = dt.strftime("%d%m%y")
                except:
                    date_str = safe_str(raw_date).replace("-", "")[:6].ljust(6, '0') # Fallback
                
            atm_id = safe_str(row[4])
            atm_id_trunc = atm_id # Do not cut off any character
            
            txn_seq = safe_str(row[5])
            file_type = safe_str(row[12])
            
            if not file_type or file_type.upper() == "N" or file_type == "None":
                continue

            type_upper = file_type.upper()
            if type_upper == "T1":
                prefix = "tffo1_disp"
            elif type_upper in ("T2", "T3"):
                prefix = "tffo2_disp"
            elif type_upper == "VC":
                prefix = "vc_disp"
            elif type_upper == "VD":
                prefix = "vd_disp"
            elif type_upper == "CR":
                prefix = "cr_disp"
            else:
                prefix = f"{type_upper}_disp"
            
            todays_date = datetime.now().strftime("%d%m%y")
            fname = f"{prefix}_{todays_date}.txt"

            # --- GENERATE VC/VD RECORD ---
            if type_upper in ("VC", "VD"):
                vc_target_date = posting_date if posting_date else raw_date
                if isinstance(vc_target_date, datetime):
                    date_str_vc = vc_target_date.strftime("%d%m%Y")
                else:
                    try:
                        dt = datetime.strptime(safe_str(vc_target_date), "%d-%b-%y")
                        date_str_vc = dt.strftime("%d%m%Y")
                    except ValueError:
                        try:
                            dt = datetime.strptime(safe_str(vc_target_date), "%d/%m/%Y")
                            date_str_vc = dt.strftime("%d%m%Y")
                        except ValueError:
                            date_str_vc = safe_str(vc_target_date).replace("-", "").replace("/", "")[:8].ljust(8, '0')

                branch_str = str(row[7]).strip() if len(row) > 7 and row[7] else "00000"
                branch_str = branch_str[-5:].zfill(5)
                
                amount_vc_paise = int(raw_amt * 1000)
                vc_year = date_str_vc[4:] if len(date_str_vc) == 8 else "2025"
                year_part = f"{vc_year}00"
                
                atm_id_6 = str(atm_id)[-6:].ljust(6, ' ') if atm_id else " "*6
                txn_seq_4 = str(txn_seq).zfill(4)[-4:] if txn_seq else "0000"
                
                dr_ac_str = str(dr_ac).strip() if dr_ac and dr_ac.lower() not in ('none', 'nan') else "0"
                cr_ac_str = str(cr_ac).strip() if cr_ac and cr_ac.lower() not in ('none', 'nan') else "0"
                
                line1 = (
                    f"02{branch_str}"
                    f"{dr_ac_str:>017}"
                    f"{cr_ac_str:>017}"
                    f"{'0'*17}"
                    f"{amount_vc_paise:017d}"
                    f"{date_str_vc}"
                    f"{' '*10}"
                    f"{year_part}"
                    f"{atm_id_6}"
                    f"{txn_seq_4}"
                    f"0P0301ATM SWITCH CENTRE DISP RDSL"
                )
                
                ref_str = str(row[0] if row[0] is not None else "").strip()
                term_bank = str(row[19] if len(row)>19 and row[19] else "SBI(MAURITIUS)LTD ATM COLLECTION ACCOUNTMU").strip()
                if not term_bank or term_bank.lower() in ('none', 'nan'):
                    term_bank = "SBI(MAURITIUS)LTD ATM COLLECTION ACCOUNTMU"
                    
                line2 = f"{ref_str:<11}{term_bank:<69}N"
                combined_record = line1 + "\n" + line2
                
                prefix_vc = "vc_disp" if type_upper == "VC" else "vd_disp"
                fname_vc = f"{prefix_vc}_{todays_date}.txt"
                if fname_vc not in files_content: files_content[fname_vc] = []
                files_content[fname_vc].append(combined_record)
                continue

            # --- GENERATE DEBIT LEG ---
            if dr_ac and dr_ac.lower() != 'none' and dr_ac.lower() != 'nan':
                 ac_type = "51"
                 if str(dr_ac).startswith("98") or str(dr_ac).startswith("48"):
                     ac_type = "54"
                 
                 left_part = f"{ac_type}{dr_ac:0>17}{amount_paise:016d}{' '*10}"
                 right_part = f"{card} {date_str} {atm_id_trunc} {txn_seq}"
                 line = (left_part + right_part).ljust(101)
                 
                 if fname not in files_content: files_content[fname] = []
                 files_content[fname].append(line)

            # --- GENERATE CREDIT LEG ---
            if cr_ac and cr_ac.lower() != 'none' and cr_ac.lower() != 'nan':
                 ac_type = "01"
                 if str(cr_ac).startswith("98") or str(cr_ac).startswith("48"):
                     ac_type = "04"
                 
                 left_part = f"{ac_type}{cr_ac:0>17}{amount_paise:016d}{' '*10}"
                 right_part = f"{card} {date_str} {atm_id_trunc} {txn_seq}"
                 line = (left_part + right_part).ljust(101)
                 
                 if fname not in files_content: files_content[fname] = []
                 files_content[fname].append(line)

        if not os.path.exists(output_dir): os.makedirs(output_dir)
        for fname, lines in files_content.items():
            with open(os.path.join(output_dir, fname), "w") as f:
                f.write("\n".join(lines))
                
    except Exception as e:
        print(f"Error processing Excel: {e}")

def process_file(args):
    inputDirectoryPath = args[0] if len(args) > 0 else ""
    outputDirectoryPath = args[1] if len(args) > 1 else ""
    
    date_suffix = datetime.now().strftime("_%d%m%y")
    default_report = f"REPORT{date_suffix}.txt"
    outputReportName = args[2] if len(args) > 2 else default_report

    report = Report()
    report.setReportName(outputReportName)
    report.setOutputDirectoryPath(outputDirectoryPath)
    report.setInputDirectoryPath(inputDirectoryPath)

    # REMOVED: Blocking isdir check. 
    # Logic is now handled by flow below.

    inputFileValidationRules = InputFileValidationRules()

    # DETECT EXCEL INPUT
    file_list = []
    
    if os.path.isfile(inputDirectoryPath):
        # We assume if it's a file, it's the Excel input based on usage
        print(f"Processing Input File: {inputDirectoryPath}")
        excel_to_cbs_files(inputDirectoryPath, outputDirectoryPath)
        # Verify the generated text files
        if os.path.exists(outputDirectoryPath):
            file_list = [
                os.path.join(outputDirectoryPath, f)
                for f in os.listdir(outputDirectoryPath)
                if os.path.isfile(os.path.join(outputDirectoryPath, f)) and f.lower().endswith('.txt') and "REPORT" not in f
            ]
    else:
        # Standard directory processing
        if not os.path.exists(inputDirectoryPath):
             print(f"ERROR: Input path does not exist: {inputDirectoryPath}")
             report.setError(True)
             report.setErrorCode("INPUT PATH DOES NOT EXIST")
             return report
             
        if not os.path.isdir(inputDirectoryPath):
             print(f"ERROR: Input is not a directory or supported file: {inputDirectoryPath}")
             report.setError(True)
             report.setErrorCode("INVALID INPUT TYPE")
             return report
        
        file_list = [
            os.path.join(inputDirectoryPath, f)
            for f in os.listdir(inputDirectoryPath)
            if os.path.isfile(os.path.join(inputDirectoryPath, f)) and f.lower().endswith('.txt')
        ]

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


def main():
    if len(sys.argv) >= 3:
        process_file(sys.argv[1:])
    else:
        from file_process_gui import main as gui_main
        gui_main()

if __name__ == "__main__":
    main()
