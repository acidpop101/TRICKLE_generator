
import os

print("--- DEBUG: Verifying main.py content ---")
path = r"d:\TRICKLE_generator\trickle_feed\main.py"
if os.path.exists(path):
    print(f"File found at: {path}")
    with open(path, 'r') as f:
        lines = f.readlines()
        # Check around line 240-250 for the fix
        for i, line in enumerate(lines):
            if "FIXED: Transaction details" in line or "Structure: 10 Spaces" in line:
                print(f"Line {i+1}: {line.strip()}")
            if "field_pad = " in line:
                print(f"Line {i+1}: {line.strip()}")
else:
    print(f"FILE NOT FOUND: {path}")
