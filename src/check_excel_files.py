
from pathlib import Path
from zipfile import is_zipfile

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = BASE_DIR / "data" / "templates"

print("===== EXCEL FILE VALIDATION =====")

for file in sorted(TEMPLATE_DIR.iterdir()):

    if not file.is_file():
        continue

    print("\nFile:", file.name)
    print("Size:", file.stat().st_size, "bytes")
    print("Extension:", file.suffix)
    print("Valid ZIP structure:", is_zipfile(file))

    # Identify the file format using its first bytes
    with open(file, "rb") as f:
        signature = f.read(8)

    print("File signature:", signature.hex())

    if is_zipfile(file):
        print("Status: ZIP format detected")
    elif signature.startswith(bytes.fromhex("D0CF11E0")):
        print("Status: Older Excel/OLE file format detected")
    elif signature.startswith(b"<") or signature.startswith(b"<!"):
        print("Status: Possible HTML/XML file")
    elif file.stat().st_size == 0:
        print("Status: Empty file")
    else:
        print("Status: Unknown or unsupported format")
