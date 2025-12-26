
# Debug script to check alignment

# 1. Check AC Details Alignment
ac_headers = ["FIID", "VOSTRO_AC", "SETTL_BGL_AC"]
ac_data = ("F001", "10983139836", "4897926042923")

print("--- AC DETAILS ALIGNMENT ---")
print(f"Headers: {len(ac_headers)}")
print(f"Data:    {len(ac_data)}")
for i, (h, d) in enumerate(zip(ac_headers, ac_data)):
    print(f"{i}: {h} -> {d}")
print("-" * 30)

# 2. Check Disputes Alignment
dispute_headers = [
    "ref", "txndate", "cardno", "acno", "atmid", "txnno", "amount", "branch",
    "debit_ac", "credit_ac", "status", "posting_date", "file_type",
    "remarks", "posting_flag", "posting_user", "card_fiid", "card_bank_name",
    "term_fiid", "term_bank_name", "comp_type", "disp_type", "src_ip", "user_name"
]

ref = "REF123"
txndate_fmt = "01-JAN-25"
cardno = "CARDNO"
acno = "ACNO"
atmid = "ATMID"
txnno = "TXNNO"
amount = "100"
branch = "00001"
debit_ac = "DR_AC"
credit_ac = "CR_AC"
file_type = "T1"
idx = 1
user_name = "USER"
card_fiid = "C_FIID"
card_bank_name = "C_BANK"
term_fiid = "T_FIID"
term_bank_name = "T_BANK"
comp_type = "SOF"
disp_type = "short"

record = [
    ref, txndate_fmt, cardno, acno, atmid, txnno, amount, branch,
    debit_ac, credit_ac, "NO", None, file_type,  f"File{idx} Entry",
    # remarks
    "NO", user_name, card_fiid, card_bank_name,
    term_fiid, term_bank_name, comp_type, disp_type, "127.0.0.1", user_name 
]

print("--- DISPUTES ALIGNMENT ---")
print(f"Headers: {len(dispute_headers)}")
print(f"Data:    {len(record)}")

if len(dispute_headers) != len(record):
    print("MISMATCH DETECTED!")
else:
    print("Counts Match.")

for i, (h, d) in enumerate(zip(dispute_headers, record)):
    print(f"{i}: {h} -> {d}")
