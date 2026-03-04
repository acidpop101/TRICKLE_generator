# TRICKLE Generator & ATM Dispute App

This project automates the processing of ATM disputes and generates "Trickle Feed" files (fixed-width text files) for the Core Banking System (CBS).

## Project Overview

The workflow consists of two main stages:
1.  **Dispute Entry**: Processing dispute details and saving them to an Excel file.
2.  **File Generation**: Converting the Excel data into specific text formats (`T1.txt`, `T2.txt`, etc.) for accounting.

## How to Use

### 1. Start the Dispute Application
Double-click **`start_app.cmd`** in the root directory.
*   **What it does**: Launches the Python Application (`ATM_DISPUTE.py`).
*   **Purpose**: Validates dispute transactions (`FOS`, `SOF`, etc.), calculates accounting legs (Credits/Debits), and saves the clean data into `data\atm_data.xlsx`.

### 2. Generate Trickle Files
Double-click **`generate_trickle.cmd`** in the root directory.
*   **What it does**: Runs the Generator script (`src\trickle_feed\main.py`).
*   **Purpose**: Reads `data\atm_data.xlsx`, enters the data into the trickle feed format (101-character fixed width), and outputs the files to `cbs_output\`.

## Directory Structure

*   **`src/`**: Source code for the applications.
    *   `atm_dispute_app/`: Logic for the dispute entry GUI/Calculations.
    *   `trickle_feed/`: Logic for generating the text files.
*   **`data/`**: Stores the database file `atm_data.xlsx`.
*   **`cbs_output/`**: The generated `.txt` files will appear here.
*   **`scripts/`**: Maintenance scripts.
