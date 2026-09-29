"""Smoke test used by the Windows CI build (see .github/workflows/build-windows.yml).

Confirms the engine module imports cleanly and the CLI can process a tiny
Thai/English .docx AND .pdf before we spend time on a full PyInstaller build.
Also confirms the bundled default technical dictionary (Kaizen/Kanban/etc.)
loads correctly.

Kept as its own script (rather than inline `python -c "..."` in the workflow
YAML) because multi-line Python literals passed as a single command-line
argument are fragile under PowerShell's argument quoting/escaping rules.
"""
import subprocess
import sys

from docx import Document

import spellcheck_engine as engine  # noqa: F401  (import check)

print("engine import OK")

if not engine.PDF_SUPPORT:
    sys.exit("pymupdf import failed -- PDF_SUPPORT is False (check requirements.txt install)")
print("pymupdf import OK")

dicts = engine.load_dictionaries(None, verbose=False)
assert "Kaizen" in dicts.custom_words, "bundled default technical dictionary did not load"
print(f"bundled default dictionary OK ({len(dicts.custom_words)} custom words loaded)")

doc = Document()
doc.add_paragraph("การทดสอบระบบ ถูกต้องครบถ้วน")
doc.save("smoke_test.docx")

subprocess.run(
    [sys.executable, "spellcheck_engine.py", "smoke_test.docx", "--quiet"],
    check=True,
)
print("CLI smoke test (docx) OK")

import pymupdf

pdoc = pymupdf.open()
page = pdoc.new_page()
page.insert_text((72, 72), "Test smoke check page for CI", fontname="helv", fontsize=14)
pdoc.save("smoke_test.pdf")
pdoc.close()

subprocess.run(
    [sys.executable, "spellcheck_engine.py", "smoke_test.pdf", "--quiet"],
    check=True,
)
print("CLI smoke test (pdf) OK")
