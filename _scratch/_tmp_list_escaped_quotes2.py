from pathlib import Path
root=Path(r"C:\Users\hercl\planning_2026")
changed=[]
for p in list(root.glob('**/*.bas'))+list(root.glob('**/*.cls'))+list(root.glob('**/*.frm')):
    try:
        text=p.read_text(encoding='latin-1')
    except Exception:
        continue
    if '\\"\\"' in text:
        # would have been changed earlier, but now list for debug
        changed.append(str(p))
print('files_with_escaped_now:', len(changed))
