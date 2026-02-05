from pathlib import Path
root=Path(r"C:\Users\hercl\planning_2026")
files=[]
for p in list(root.glob('**/*.bas'))+list(root.glob('**/*.cls'))+list(root.glob('**/*.frm')):
    try:
        text=p.read_text(encoding='latin-1')
    except Exception:
        continue
    if '\\"\\"' in text:
        files.append(str(p))
print('\n'.join(files))
