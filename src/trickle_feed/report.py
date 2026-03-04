
from typing import List
from file_object import FileObject

class Report:
    DOUBLE_LINE = "===================================================================================================================================================="
    HEADER = "\n".join([
        DOUBLE_LINE,
        "{:<80} {:>18} {:>16} {:>16} {:>16}".format("File Name", "Record Count", "Dr Amount", "Cr Amount", "CHECK"),
        DOUBLE_LINE
    ])

    def __init__(self):
        self.reportString = None
        self.reportName = "REPORT.txt"
        self.outputDirectoryPath = ""
        self.inputDirectoryPath = ""
        self.isError = False
        self.sErrorCode = None
        self.fileEntries: List[FileObject] = []

    def getReportHeader(self): return self.HEADER

    def setReportContent(self, report):
        import datetime
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        total_records = 0
        total_dr = 0.0
        total_cr = 0.0
        
        lines = [
            "", 
            self.DOUBLE_LINE,
            f"                                                 TRICKLEFEED CR/DR REPORT ({current_time})                                                                  ",
            self.getReportHeader()
        ]

        for fileRecord in report.getFileEntries():
            if fileRecord.isErrorFlag():
                rec_count_col = fileRecord.getErrorCode() or ""
                dr_amt_col = cr_amt_col = check_col = ""
            else:
                rec_count = fileRecord.getRecordCount() or 0
                dr_amt = fileRecord.getDebitAmount() or 0.0
                cr_amt = fileRecord.getCreditAmount() or 0.0
                
                # Accumulate totals
                total_records += rec_count
                total_dr += dr_amt
                total_cr += cr_amt
                
                rec_count_col = f"{rec_count:,}"
                dr_amt_col = f"{dr_amt:,.2f}"
                cr_amt_col = f"{cr_amt:,.2f}"
                check_col = f"{(cr_amt - dr_amt):,.2f}"

            lines.append("{:<80}{:>18}{:>16}{:>16}{:>16}".format(
                fileRecord.getFileName() or "", rec_count_col, dr_amt_col, cr_amt_col, check_col))

        lines.append(self.DOUBLE_LINE)
        
        # Add totals row
        lines.append("{:<80}{:>18}{:>16}{:>16}{:>16}".format(
            "TOTALS", 
            f"{total_records:,}", 
            f"{total_dr:,.2f}", 
            f"{total_cr:,.2f}", 
            f"{(total_cr - total_dr):,.2f}"
        ))
        lines.append(self.DOUBLE_LINE)
        
        self.setReport("\n".join(lines))

    def getFileEntries(self): return self.fileEntries
    def setFileEntries(self, entries): self.fileEntries = entries
    def getErrorCode(self): return self.sErrorCode
    def setErrorCode(self, code): self.sErrorCode = code
    def isErrorFlag(self): return self.isError
    def setError(self, flag): self.isError = flag
    def getInputDirectoryPath(self): return self.inputDirectoryPath
    def setInputDirectoryPath(self, path): self.inputDirectoryPath = path
    def getOutputDirectoryPath(self): return self.outputDirectoryPath
    def setOutputDirectoryPath(self, path): self.outputDirectoryPath = path
    def getReportName(self): return self.reportName
    def setReportName(self, name): self.reportName = name
    def getReport(self): return self.reportString
    def setReport(self, report_str): self.reportString = report_str
