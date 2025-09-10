import py_compile
from pathlib import Path

p = Path(r"e:/ea-cli/easy_access")
failed = False
for f in sorted(p.rglob("*.py")):
    try:
        py_compile.compile(str(f), doraise=True)
    except Exception as e:
        print("FAIL", f, e)
        failed = True
if not failed:
    print("All files compiled successfully")
