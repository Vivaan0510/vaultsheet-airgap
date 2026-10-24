import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color, HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Register Windows Nirmala UI / Mangal fallback
FONT_DEVA = "Helvetica"
for candidate in [r"C:\Windows\Fonts\Nirmala.ttf", r"C:\Windows\Fonts\mangal.ttf", r"C:\Windows\Fonts\arialuni.ttf"]:
    if os.path.exists(candidate):
        try:
            pdfmetrics.registerFont(TTFont("WinDeva", candidate))
            FONT_DEVA = "WinDeva"
            break
        except Exception:
            pass

def generate_all_test_pdfs():
    # Wrapper to satisfy app.py import
    pass

if __name__ == "__main__":
    generate_all_test_pdfs()
