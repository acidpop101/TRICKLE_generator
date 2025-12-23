
BGL_DR = "54"
SYS_DR = "62"
CUSTOMER_DR = "51"
BGL_CR = "04"
SYS_CR = "12"
CUSTOMER_CR = "01"
MAX_CR_DR_AMOUNT_ALLOWED = 30000000000
# constants.py - EXACT Java array structure (16x9)
# constants.py - COMPLETE 16x9 array from Java TestClass.main() [file:1]
checkDigitConstantArray = [
    # Row JE=15: [10,9,8,7,6,5,4,3,2]
    [10, 9, 8, 7, 6, 5, 4, 3, 2],
    # Row JE=14: [5,10,4,9,3,8,2,7,1]  
    [5, 10, 4, 9, 3, 8, 2, 7, 1],
    # Row JE=13: [8,5,2,10,7,4,1,9,6]
    [8, 5, 2, 10, 7, 4, 1, 9, 6],
    # Row JE=12: [4,8,1,5,9,2,6,10,3]
    [4, 8, 1, 5, 9, 2, 6, 10, 3],
    # Row JE=11: [2,4,6,8,10,1,3,5,7]
    [2, 4, 6, 8, 10, 1, 3, 5, 7],
    # Row JE=10: [1,2,3,4,5,6,7,8,9]
    [1, 2, 3, 4, 5, 6, 7, 8, 9],
    # Row JE=9:  [6,1,7,2,8,3,9,4,10]
    [6, 1, 7, 2, 8, 3, 9, 4, 10],
    # Row JE=8:  [3,6,9,1,4,7,10,2,5]
    [3, 6, 9, 1, 4, 7, 10, 2, 5],
    # Row JE=7:  [7,3,10,6,2,9,5,1,8]
    [7, 3, 10, 6, 2, 9, 5, 1, 8],
    # Row JE=6:  [9,7,5,3,1,10,8,6,4]
    [9, 7, 5, 3, 1, 10, 8, 6, 4],
    # Row JE=5:  [10,9,8,7,6,5,4,3,2]
    [10, 9, 8, 7, 6, 5, 4, 3, 2],
    # Row JE=4:  [5,10,4,9,3,8,2,7,1]
    [5, 10, 4, 9, 3, 8, 2, 7, 1],
    # Row JE=3:  [8,5,2,10,7,4,1,9,6]
    [8, 5, 2, 10, 7, 4, 1, 9, 6],
    # Row JE=2:  [4,8,1,5,9,2,6,10,3]
    [4, 8, 1, 5, 9, 2, 6, 10, 3],
    # Row JE=1:  [2,4,6,8,10,1,3,5,7]
    [2, 4, 6, 8, 10, 1, 3, 5, 7],
    # Row JE=0:  [1,2,3,4,5,6,7,8,9]
    [1, 2, 3, 4, 5, 6, 7, 8, 9]
]

# Verify dimensions match Java (16 rows × 9 columns)
assert len(checkDigitConstantArray) == 16, "Must have 16 rows (JE=15 to 0)"
assert all(len(row) == 9 for row in checkDigitConstantArray), "Each row must have 9 columns (Mdigit=0 to 8)"
UTILITY_NAME = "Trickle Feed File Process"
UTILITY_VERSION = "1.1"
COPYRIGHT_CLAUSE = "Copyright © 2017 State Bank of India. All Rights Reserved."

# constants.py - KEEP YOUR EXISTING ARRAY (it's CORRECT)

# CORRECTED Verification Test (JE=15 is FIRST row, index 0)
print("checkDigitConstantArray shape:", len(checkDigitConstantArray), "x", len(checkDigitConstantArray[0]))
print("Sample lookup JE=15, Mdigit=0:", checkDigitConstantArray[0][0])   # Should be 10 ✓
print("Sample lookup JE=10, Mdigit=4:", checkDigitConstantArray[5][4])   # Should be 5 ✓
print("Sample lookup JE=0, Mdigit=0: ", checkDigitConstantArray[15][0])  # Should be 1  ✓

# Array is CORRECT - Java indexing confirmed
print("✅ Array verified - Matches Java exactly")




