@echo off
TITLE VaultSheet - 100% Offline Forensic Accounting Workstation
cd /d "C:\Users\Vivaan Kesarwani\vaultsheet"
call venv\Scripts\activate.bat
if not exist test_statements (
    python generate_test_pdfs.py
)
streamlit run app.py --browser.gatherUsageStats=false --server.address=127.0.0.1
pause
