#!/bin/bash
cd "C:\Users\Vivaan Kesarwani\vaultsheet"
source venv/bin/activate
if [ ! -d "test_statements" ]; then
    python generate_test_pdfs.py
fi
streamlit run app.py --browser.gatherUsageStats=false --server.address=127.0.0.1
