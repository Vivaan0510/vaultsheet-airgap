import os
import sys
from pathlib import Path
import streamlit.web.cli as stcli


def resolve_path(relative_path: str) -> str:
    """Resolve resource path for both local dev and PyInstaller frozen .exe."""
    base_path = getattr(sys, "_MEIPASS", Path(__file__).parent.resolve())
    return str(Path(base_path) / relative_path)


if __name__ == "__main__":
    # Ensure working directory is set to the executable's directory
    if getattr(sys, "frozen", False):
        os.chdir(Path(sys.executable).parent)
    else:
        os.chdir(Path(__file__).parent.resolve())

    app_path = resolve_path("app.py")

    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
        "--server.address=127.0.0.1",
        "--client.toolbarMode=minimal",
    ]
    sys.exit(stcli.main())