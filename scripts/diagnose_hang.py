"""Run after pytest finishes (or when pytest appears to hang) to list live threads
and pending asyncio tasks. Useful to paste back for diagnosis.

Usage (PowerShell):
    python .\\scripts\\diagnose_hang.py

It prints thread names and repr() of pending asyncio tasks.
"""

import asyncio
import threading

print("=== Active threads ===")
for t in threading.enumerate():
    print(f"- {t.name} (daemon={t.daemon})")

print("\n=== Pending asyncio tasks (in current loop) ===")
try:
    loop = asyncio.get_event_loop()
    tasks = asyncio.all_tasks(loop)
    if tasks:
        for t in tasks:
            print(f"- {t!r}")
    else:
        print("(no pending tasks)")
except Exception as e:
    print("Could not inspect asyncio tasks:", e)

print("\n=== End diagnostics ===")
