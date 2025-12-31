
# 16x9 Matrix from constants.py
checkDigitConstantArray = [
    [10, 9, 8, 7, 6, 5, 4, 3, 2],
    [5, 10, 4, 9, 3, 8, 2, 7, 1],
    [8, 5, 2, 10, 7, 4, 1, 9, 6],
    [4, 8, 1, 5, 9, 2, 6, 10, 3],
    [2, 4, 6, 8, 10, 1, 3, 5, 7],
    [1, 2, 3, 4, 5, 6, 7, 8, 9],
    [6, 1, 7, 2, 8, 3, 9, 4, 10],
    [3, 6, 9, 1, 4, 7, 10, 2, 5],
    [7, 3, 10, 6, 2, 9, 5, 1, 8],
    [9, 7, 5, 3, 1, 10, 8, 6, 4],
    [10, 9, 8, 7, 6, 5, 4, 3, 2],
    [5, 10, 4, 9, 3, 8, 2, 7, 1],
    [8, 5, 2, 10, 7, 4, 1, 9, 6],
    [4, 8, 1, 5, 9, 2, 6, 10, 3],
    [2, 4, 6, 8, 10, 1, 3, 5, 7],
    [1, 2, 3, 4, 5, 6, 7, 8, 9]
]

def get_check_digit(acc_num_str):
    """
    Port of input_file_validation_rules.py logic
    """
    try:
        Macno = int(acc_num_str)
    except:
        return "?"

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

test_cases = [
    ("9858100138", 8),
    ("9858100177", 4),
    ("9858120134", 4)
]

print(f"{'Account':<12} | {'Exp':<3} | {'Calc':<4} | {'Match':<5}")
print("-" * 40)
for acc, expected in test_cases:
    calc = get_check_digit(acc)
    match = int(calc) == expected
    print(f"{acc:<12} | {expected:<3} | {calc:<4} | {match:<5}")
