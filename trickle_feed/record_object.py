
class RecordObject:
    def __init__(self):
        self.sRecord = None
        self.sErrorCode = None
        self.isError = False
        self.sAccountType = None
        self.lAmount = None
        self.lAccountNumber = None

    def getAccountNumber(self): return self.lAccountNumber
    def setAccountNumber(self, val): self.lAccountNumber = val
    def getAccountType(self): return self.sAccountType
    def setAccountType(self, val): self.sAccountType = val
    def getAmount(self): return self.lAmount
    def setAmount(self, val): self.lAmount = val
    def isErrorFlag(self): return self.isError
    def setError(self, flag): self.isError = flag
    def getErrorCode(self): return self.sErrorCode
    def setErrorCode(self, code): self.sErrorCode = code
    def getRecord(self): return self.sRecord
    def setRecord(self, rec): self.sRecord = rec
    def setRecordToInitial(self):
        self.sRecord = self.sErrorCode = None
        self.isError = False
        self.sAccountType = self.lAmount = self.lAccountNumber = None
