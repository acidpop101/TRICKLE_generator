
line = "54000000985810013880000000000100000          4359786000300866 202520 000138136 2140                  "
print(f"Total Length: {len(line)}")
print("0123456789" * 10)
print(line)

print("\nPossible Fields:")
print(f"Type (0-2): '{line[0:2]}'")
print(f"Acct (2-19): '{line[2:19]}'")
print(f"Amt  (19-35): '{line[19:35]}'") 
print(f"Rest (35-):   '{line[35:]}'")

# Split rest
rest = line[35:]
print(f"\nRest Analysis ({len(rest)} chars):")
print(f"'{rest}'")
