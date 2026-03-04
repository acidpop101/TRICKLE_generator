# Technical Specifications

## 1. Trickle Feed File Format
The Core Banking System (CBS) expects a fixed-width text file where each line represents a transaction leg.
**Total Line Length**: 101 Characters.

| Position (0-Index) | Length | Field Description | Format / Notes |
| :--- | :--- | :--- | :--- |
| **00 - 01** | 2 | **Account Type** | Code denoting the nature of the account (see Section 3). |
| **02 - 18** | 17 | **Account Number** | Right-aligned. |
| **19 - 34** | 16 | **Amount** | In **paise** (lowest currency unit). Zero-padded (e.g., `0000000000010000` = ₹100.00). |
| **35 - 44** | 10 | **Filler** | Empty spaces. |
| **45 - 60** | 16 | **Card Number** | The 16-digit ATM card number. |
| **61 - 61** | 1 | **Space** | Separator. |
| **62 - 67** | 6 | **Txn Date** | Format: `DDMMYY` (e.g., `010125`). |
| **68 - 68** | 1 | **Space** | Separator. |
| **69 - 77** | 9 | **ATM ID** | Last 9 digits of the Terminal ID. |
| **78 - 78** | 1 | **Space** | Separator. |
| **79 - 82** | 4 | **Txn Sequence** | Last 4 digits of the System Trace Audit Number (STAN). |
| **83 - 100** | 18 | **Filler** | Remaining spaces to complete the record. |

---

## 2. Dispute Types
These codes classify the reason for the ATM dispute.

*   **`short` (Short Cash / Partial Dispense)**
    *   **Meaning**: The customer requested an amount (e.g., 5000) but received less (e.g., 2000) or nothing, yet the account was debited.
    *   **Action**: Determine the shortage amount and refund the customer.

*   **`dd` (Duplicate Debit)**
    *   **Meaning**: The customer's account was debited **twice** for a single transaction attempt.
    *   **Action**: Reverse one of the debits.
    *   *Note*: For FOS (Other Bank) transactions, this involves crediting the Issuer Bank's Vostro account.

*   **`unsucc` (Unsuccessful Transaction)**
    *   **Meaning**: The transaction failed at the ATM (no cash dispensed), but the account was still debited.
    *   **Action**: Full reversal of the debited amount.

*   **`full` (Full Reversal)**
    *   **Meaning**: Similar to `unsucc`, explicitly indicating the entire transaction amount needs valid reversal.

---

## 3. Account Type Codes
The first 2 characters of the trickle feed line indicate the account category.

*   **`01`**: **Customer Credit** (Refund to a Customer's Savings/Current Account).
*   **`51`**: **Customer Debit** (Deduction from a Customer's Account).
*   **`12`**: **System Credit** (Credit to an Internal Office Account, e.g., Vostro, Suspense, or GL).
*   **`62`**: **System Debit** (Debit from an Internal Office Account).
*   **`04`**: **BGL Credit** (Branch General Ledger Credit - specific internal use).
*   **`54`**: **BGL Debit** (Branch General Ledger Debit - specific internal use).

---

## 4. Transaction Categories

*   **SOF (Financial On Us)**: Our Card used at Our ATM.
*   **FOS (Financial Other Switch)**: Other Bank's Card used at Our ATM. (Requires Settlement/Vostro adjustment).
*   **XYZ (Grid/Network)**: Transactions routed through specific payment networks (e.g., NFS, Visa).
