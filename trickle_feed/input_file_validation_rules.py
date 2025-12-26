
from constants import BGL_DR, SYS_DR, CUSTOMER_DR, BGL_CR, SYS_CR, CUSTOMER_CR, checkDigitConstantArray, MAX_CR_DR_AMOUNT_ALLOWED
import error_code as ec

class InputFileValidationRules:
    def inputFileRules(self, record, file_obj):
        self.checkRecordLength(record, file_obj)
        if not record.isErrorFlag():
            self.setRecordDetails(record, file_obj)
        if not record.isErrorFlag():
            self.checkAccountNumber(record, file_obj)
        if not record.isErrorFlag():
            self.calculateAmount(record, file_obj)

    def checkRecordLength(self, record, file_obj):
        rec_len = len(record.getRecord())
        if self.getRecordLength() < rec_len:
            record.setError(True)
            record.setErrorCode(ec.LENGTH_ERR)
        elif self.getAccountNumberAndAmountLength() > rec_len:
            record.setError(True)
            record.setErrorCode(ec.LENGTH_ACCNT_AMT_ERR)

    def calculateAmount(self, record, file_obj):
        acc_type = record.getAccountType()
        amt = record.getAmount() or 0
        if acc_type and acc_type.upper() in (BGL_DR, SYS_DR, CUSTOMER_DR):
            file_obj.setDebitAmount(file_obj.getDebitAmount() + amt)
        if acc_type and acc_type.upper() in (BGL_CR, SYS_CR, CUSTOMER_CR):
            file_obj.setCreditAmount(file_obj.getCreditAmount() + amt)

    def setFileObjectWithRecordDetails(self, record, file_obj):
        file_obj.setError(record.isErrorFlag())
        file_obj.setErrorCode(record.getErrorCode())
        if not record.isErrorFlag():
            message1 = ""
            if file_obj.getDebitAmount() // 100 > MAX_CR_DR_AMOUNT_ALLOWED:
                file_obj.setError(True)
                message1 = "CREDIT "
            if file_obj.getCreditAmount() // 100 > MAX_CR_DR_AMOUNT_ALLOWED:
                file_obj.setError(True)
                message1 = message1 + (", " if file_obj.isErrorFlag() else "") + "DEBIT "
            if file_obj.isErrorFlag():
                file_obj.setErrorCode(message1 + ec.MAX_CR_DR_AMOUNT_ERR)

    def checkAccountNumber(self, record, fileobj):
        # SKIP validation for BGL/System accounts which can be alphanumeric
        acc_type = record.getAccountType()
        if acc_type in (BGL_DR, SYS_DR, BGL_CR, SYS_CR):
            return

        if record.getAccountNumber() == 0:
            record.setError(True)
            record.setErrorCode(ec.ACCOUNT_NUMBER_ZERO_ERR)
            return
    
        account_str = str(record.getAccountNumber())
        account_length = len(account_str)
    
        if not (11 <= account_length <= 13):
            record.setError(True)
            record.setErrorCode(ec.ACCOUNTNUMBERCHARSERR)
            return
    
        last_digit = record.getAccountNumber() % 10
        calculated_check_digit = self.getCheckDigitNumber(record.getAccountNumber() // 10)
    
        if last_digit != calculated_check_digit:
            record.setError(True)
            record.setErrorCode(ec.ACCOUNTCHKDIGITERR)


        acc_no_str = str(acc_no)
        length = len(acc_no_str)
        if 11 <= length <= 13:
            last_digit = acc_no % 10
            base_no = acc_no // 10
            check_digit = self.getCheckDigitNumber(base_no)
            if last_digit != check_digit:
                record.setError(True)
                record.setErrorCode(ec.ACCOUNT_CHK_DIGIT_ERR)
        else:
            record.setError(True)
            record.setErrorCode(ec.ACCOUNT_NUMBER_CHARS_ERR)

    def getRecordLength(self): return ec.LENGTH_RECORD
    def getAccountNumberAndAmountLength(self): return ec.LENGTH_ACCNT_AMT

    def setRecordDetails(self, record, file_obj):
        rec = record.getRecord()
        sAccountType = rec[0:2]
        sAccount = rec[2:19]
        sAmount = rec[19:35]
        print(f"DEBUG: Processing Line: Type='{sAccountType}', Account='{sAccount}'") # Debug print
        record.setAccountType(sAccountType)
        record.setAccountType(sAccountType)
        
        # CHANGED: Only enforce int cast for Customer accounts. Allow string for others.
        if sAccountType in (BGL_DR, SYS_DR, BGL_CR, SYS_CR):
             record.setAccountNumber(sAccount.strip()) # Store as string
        else:
            try:
                record.setAccountNumber(int(sAccount))
            except:
                record.setError(True)
                record.setErrorCode(ec.ACCOUNT_NUMERIC_ERR)
                return
        
        try:
            record.setAmount(int(sAmount))
        except:
            record.setError(True)
            record.setErrorCode(ec.AMOUNT_NUMERIC_ERR)

    def isErrorFlag(self): return self.isError

    def getCheckDigitNumber(self, account_number):
        """Exact Java algorithm port"""
        Macno = account_number
        JE = 15
        Mchkdigit = 0
        Cdigit = 0
    
        while Macno > 0:
            iLastDigit = Macno % 10
            if iLastDigit > 0:  # Skip 0 digits (Java logic)
                Mdigit = iLastDigit - 1  # 1→0, 2→1, ..., 9→8
                indexI = JE
                indexJ = Mdigit
                # CRITICAL: Exact bounds checking
                if 0 <= indexI < 16 and 0 <= indexJ < 9:
                    Mchkdigit += Constants.checkDigitConstantArray[indexI][indexJ]
        
            Macno //= 10
            JE -= 1  # Decrements from 15→0
    
        Cdigit = Mchkdigit % 10
        return Cdigit

   
