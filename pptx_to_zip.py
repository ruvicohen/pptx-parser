import zipfile
from pathlib import Path

pptx_path = "test.pptx"
out_dir = Path("pptx_unzipped")

out_dir.mkdir(exist_ok=True)

with zipfile.ZipFile(pptx_path, "r") as z:
    z.extractall(out_dir)

print("Extracted to:", out_dir.resolve())