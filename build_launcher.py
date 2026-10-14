import os
import platform
import subprocess
import sys
from pathlib import Path


def create_shell_launchers():
    """Generates 1-click double-click scripts for Windows and macOS/Linux."""
    root_dir = Path(__file__).parent.resolve()

    # Windows batch script (.bat)
    bat_content = f"""@echo off
TITLE VaultSheet - 100% Offline Forensic Accounting Workstation
cd /d "{root_dir}"
call venv\\Scripts\\activate.bat
if not exist test_statements (
    python generate_test_pdfs.py
)
streamlit run app.py --browser.gatherUsageStats=false --server.address=127.0.0.1
pause
"""
    bat_path = root_dir / "start_vaultsheet.bat"
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(bat_content)

    # Mac/Linux shell script (.sh)
    sh_content = f"""#!/bin/bash
cd "{root_dir}"
source venv/bin/activate
if [ ! -d "test_statements" ]; then
    python generate_test_pdfs.py
fi
streamlit run app.py --browser.gatherUsageStats=false --server.address=127.0.0.1
"""
    sh_path = root_dir / "start_vaultsheet.sh"
    with open(sh_path, "w", encoding="utf-8") as f:
        f.write(sh_content)
    try:
        os.chmod(sh_path, 0o755)
    except Exception:
        pass

    print(f"Created Windows launcher: {bat_path}")
    print(f"Created Mac/Linux launcher: {sh_path}")


def run_pyinstaller_bundle():
    """Generates the single-executable build specification for PyInstaller."""
    root_dir = Path(__file__).parent.resolve()
    entry_script = root_dir / "run_vaultsheet.py"

    runner_code = """import os
import sys
import streamlit.web.cli as stcli

if __name__ == '__main__':
    sys.argv = [
        "streamlit",
        "run",
        "app.py",
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
        "--server.address=127.0.0.1"
    ]
    sys.exit(stcli.main())
"""
    with open(entry_script, "w", encoding="utf-8") as f:
        f.write(runner_code)

    print(f"Created PyInstaller entrypoint: {entry_script}")
    print("\nTo build a standalone .exe distribution, run:")
    print(
        "pyinstaller --noconfirm --onedir --windowed "
        "--copy-metadata streamlit "
        "--collect-all pdfplumber "
        "--collect-all rapidocr_onnxruntime "
        "--collect-all pypdfium2 "
        "--add-data test_statements;test_statements "
        "run_vaultsheet.py"
    )


if __name__ == "__main__":
    create_shell_launchers()
    run_pyinstaller_bundle()