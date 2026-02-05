from pathlib import Path
import re
p = Path(r"C:\Users\hercl\planning_2026\UserForm1.frm")
text = p.read_text(encoding="latin-1")
if "AfficherMasquerLignesDynamiques" in text:
    print("already")
    raise SystemExit

m = re.search(r"Sub\s+AfficherMasquerLignes\(\).*?End Sub", text, flags=re.S)
if not m:
    raise SystemExit("AfficherMasquerLignes block not found")
block = m.group(0)
stub = block + "\n\nPrivate Sub AfficherMasquerLignesDynamiques()\n    ' TODO: remplacer par la vraie logique si besoin\n    AfficherMasquerLignes\nEnd Sub"
text = text.replace(block, stub, 1)
p.write_text(text, encoding="latin-1")
print("inserted")
