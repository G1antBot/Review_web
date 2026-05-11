import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

DOCX_PATH = Path(r"I:\大学\cat\诊断学考试简答及论述.docx")
OUT_PATH = Path(r"I:\大学\cat\untitled1\docs\diagnostics_docx_lines.txt")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

with zipfile.ZipFile(DOCX_PATH, "r") as zf:
    xml_bytes = zf.read("word/document.xml")

root = ET.fromstring(xml_bytes)
lines = []
for p in root.findall(".//w:p", NS):
    texts = [t.text or "" for t in p.findall(".//w:t", NS)]
    line = "".join(texts).strip()
    if line:
        lines.append(line)

OUT_PATH.write_text("\n".join(lines), encoding="utf-8")
print(f"导出段落 {len(lines)} 行 -> {OUT_PATH}")
for i, line in enumerate(lines[:120], 1):
    print(f"{i:03d}: {line}")
