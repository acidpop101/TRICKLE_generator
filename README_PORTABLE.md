# Portable Project Guide

This project is now designed to be portable. You can move the entire folder to any other Windows machine, and it should run essentially "out of the box", provided Python is installed.

## Prerequisites
- **Python 3.7+**: Must be installed on the target machine and available in the system PATH.
  - To verify, open cmd and type `python --version`.

## How to Move/Deploy
1. Zip the entire project folder.
2. Unzip it on the new machine (e.g., `D:\MyProject` or `C:\Users\Name\Desktop\Project`).
3. Ensure no files are missing (specially `data/atm_data.xlsx` and `requirements.txt`).

## Running the Application
### 1. ATM Dispute App (GUI)
- Double-click **`start_app.cmd`**.
- It will automatically:
  - Check if Python is available.
  - Check if required libraries (like `openpyxl`) are installed.
  - **Install missing dependencies** automatically from `requirements.txt`.
  - Launch the application.

### 2. Trickle Generator (Script)
- Double-click **`generate_trickle.cmd`**.
- It performs the same checks and runs the generation logic.

## Directory Structure
- **`src/`**: Source code (do not modify manually).
- **`data/`**: Input Excel file (`atm_data.xlsx`).
- **`cbs_output/`**: Generated text files will appear here.
- **`requirements.txt`**: List of dependencies.
