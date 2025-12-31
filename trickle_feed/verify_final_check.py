
import os

FILE_PATH = r"d:\TRICKLE_generator\trickle_feed\cbs_output\T1_301225.dat"

def verify_file():
    if not os.path.exists(FILE_PATH):
        print(f"File not found: {FILE_PATH}")
        return

    print(f"{'Row':<5} | {'Account (Pos 2-19)':<20} | {'Check Suffix':<3} | {'Expected'}")
    print("-" * 55)
    
    with open(FILE_PATH, 'r') as f:
        # Read debit lines (Lines 1, 3, 5 = Indices 0, 2, 4)
        lines = f.readlines()
        
    expected_checks = [4, 4, 8] # 00177->4, 20134->4, 00138->8
    
    debit_indices = [0, 2, 4]
    
    for i, idx in enumerate(debit_indices):
        if idx >= len(lines): break
        line = lines[idx]
        acct = line[2:19]
        chk = acct[-1]
        exp = expected_checks[i]
        
        match = "MATCH" if str(chk) == str(exp) else "FAIL"
        print(f"{i+1:<5} | {acct:<20} | {chk:<12} | {exp:<3} ({match})")
        
        print(f"      Line: {line.strip()}")

if __name__ == "__main__":
    verify_file()
