import os
from pathlib import Path
from PIL import Image

# Seek the downloaded image
downloads_path = Path.home() / "Downloads"
source_image = downloads_path / "Secure Finance Dashboard Icon.png"

# Fallback to local if not in Downloads
if not source_image.exists():
    source_image = Path("./Secure Finance Dashboard Icon.png")

target_ico = Path("./vaultsheet_icon.ico")

if source_image.exists():
    print(f"Found source image: {source_image}")
    try:
        # Open the image and convert it to a standard Windows icon
        img = Image.open(source_image)
        # Windows icons contain several standard downscaled sizes (e.g., 256, 128, 64, 48, 32, 16)
        sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
        img.save(target_ico, format="ICO", sizes=sizes)
        print(f"Successfully created: {target_ico.resolve()}")
    except Exception as e:
        print(f"Error converting image: {e}")
else:
    print(f"Could not find 'Secure Finance Dashboard Icon.png' in either your Downloads or project folder.")
    print("Please make sure the file name is correct and in one of those spots!")