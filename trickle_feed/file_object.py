
class FileObject:
    def __init__(self):
        self.lCreditAmount = 0
        self.lDebitAmount = 0
        self.lRecordCount = 0
        self.sFileName = None
        self.sErrorCode = None
        self.isError = False

    def getCreditAmount(self): return self.lCreditAmount
    def setCreditAmount(self, val): self.lCreditAmount = val
    def getDebitAmount(self): return self.lDebitAmount
    def setDebitAmount(self, val): self.lDebitAmount = val
    def getRecordCount(self): return self.lRecordCount
    def setRecordCount(self, val): self.lRecordCount = val
    def getFileName(self): return self.sFileName
    def setFileName(self, name): self.sFileName = name
    def getErrorCode(self): return self.sErrorCode
    def setErrorCode(self, code): self.sErrorCode = code
    def isErrorFlag(self): return self.isError
    def setError(self, flag): self.isError = flag
