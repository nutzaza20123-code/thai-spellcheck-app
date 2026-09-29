"""Smoke test used by the Windows CI build (see .github/workflows/build-windows.yml).

Confirms the engine module imports cleanly and the CLI can process a tiny
Thai/English .docx before we spend time on a full PyInstaller build.

Kept as its own script (rather than inline `python -c "..."` in the workflow
YAML) because multi-line Python literals passed as a single command-line
argument are fragile under PowerShell's argument quoting/escaping rules.
"""
import subprocess
import sys

from docx import Document

import spellcheck_engine  # noqa: F401  (import check)

print("engine import OK")

doc = Document()
doc.add_paragraph("การทดสอบระบบ ถูกต้องครบถ้วน")
doc.save("smoke_test.docx")

subprocess.run(
    [sys.executable, "spellcheck_engine.py", "smoke_test.docx", "--quiet"],
    check=True,
)

print("CLI smoke test OK")
