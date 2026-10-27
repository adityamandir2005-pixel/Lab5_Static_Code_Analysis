"""Build the Lab 5 static code analysis report (DOCX).

Run from the lab directory:
    python build_report.py
"""

import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

LAB = Path(__file__).resolve().parent
REPORTS = LAB / "reports"

DARK_BLUE = RGBColor(0x1F, 0x38, 0x64)
ACCENT = RGBColor(0x2E, 0x74, 0xB5)
GREY = RGBColor(0x59, 0x59, 0x59)
HEADER_FILL = "1F3864"
ZEBRA_FILL = "F2F5FA"
CODE_FILL = "F4F4F4"


# --------------------------------------------------------------------------
# report parsing helpers
# --------------------------------------------------------------------------
def read_text(path):
    """Return the report file as text (missing/empty files return "")."""
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8", errors="replace")


def pylint_messages(path):
    """Return [(line, code)] for a pylint text report."""
    pattern = re.compile(r"^user_utils\.py:(\d+):\d+: ([A-Z]\d{4}):")
    found = []
    for raw in read_text(path).splitlines():
        match = pattern.match(raw)
        if match:
            found.append((int(match.group(1)), match.group(2)))
    return found


def pylint_score(path):
    """Return the 'rated at X/10' score from a pylint report."""
    match = re.search(r"rated at ([\d.]+)/10", read_text(path))
    return match.group(1) if match else "n/a"


def flake8_messages(path):
    """Return [(line, code)] for a flake8 text report."""
    pattern = re.compile(r"^user_utils\.py:(\d+):\d+: ([A-Z]\d+) ")
    found = []
    for raw in read_text(path).splitlines():
        match = pattern.match(raw)
        if match:
            found.append((int(match.group(1)), match.group(2)))
    return found


def bandit_issues(path):
    """Return [(line, code, severity)] for a bandit text report."""
    text = read_text(path)
    found = []
    for block in text.split(">> Issue:")[1:]:
        code = re.search(r"\[(B\d+):", block)
        severity = re.search(r"Severity:\s+(\w+)", block)
        location = re.search(r"Location:\s+\.\\?user_utils\.py:(\d+):", block)
        if code and severity and location:
            found.append((int(location.group(1)), code.group(1), severity.group(1)))
    return found


def codes_to_lines(messages, wanted):
    """Return sorted line numbers whose code is in `wanted`."""
    wanted = set(wanted)
    return sorted({line for line, code in messages if code in wanted})


def fmt_lines(numbers):
    """Render line numbers as a compact, human readable string."""
    if not numbers:
        return "-"
    parts = []
    start = prev = numbers[0]
    for number in numbers[1:]:
        if number == prev + 1:
            prev = number
            continue
        parts.append(str(start) if start == prev else f"{start}-{prev}")
        start = prev = number
    parts.append(str(start) if start == prev else f"{start}-{prev}")
    return ", ".join(parts)


# --------------------------------------------------------------------------
# docx helpers
# --------------------------------------------------------------------------
def shade(cell, fill):
    """Fill a table cell background with a solid colour."""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def cell_text(cell, text, bold=False, size=8.0, color=None, align=None):
    """Set a cell to one formatted run, without extra paragraph spacing."""
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_before = Pt(1)
    paragraph.paragraph_format.space_after = Pt(1)
    paragraph.paragraph_format.line_spacing = 1.0
    if align is not None:
        paragraph.alignment = align
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Calibri"
    if color is not None:
        run.font.color.rgb = color


def repeat_header(row):
    """Mark a table row as a repeating header."""
    tr_pr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    tr_pr.append(header)


def no_split(row):
    """Prevent a row from splitting across pages."""
    tr_pr = row._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:cantSplit"))


def add_table(doc, headers, rows, widths, font_size=8.0, zebra=True):
    """Create a formatted table with a coloured, repeating header."""
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    header_row = table.rows[0]
    repeat_header(header_row)
    no_split(header_row)
    for index, (title, width) in enumerate(zip(headers, widths)):
        cell = header_row.cells[index]
        cell_text(cell, title, bold=True, size=font_size,
                  color=RGBColor(0xFF, 0xFF, 0xFF),
                  align=WD_ALIGN_PARAGRAPH.CENTER)
        shade(cell, HEADER_FILL)
        cell.width = Cm(width)

    for row_index, values in enumerate(rows):
        row = table.add_row()
        no_split(row)
        for col_index, (value, width) in enumerate(zip(values, widths)):
            cell = row.cells[col_index]
            cell_text(cell, value, size=font_size)
            cell.width = Cm(width)
            if zebra and row_index % 2 == 1:
                shade(cell, ZEBRA_FILL)

    for index, width in enumerate(widths):
        table.columns[index].width = Cm(width)
    return table


def caption(doc, text):
    """Add a small italic caption below a table."""
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(2)
    paragraph.paragraph_format.space_after = Pt(10)
    run = paragraph.add_run(text)
    run.italic = True
    run.font.size = Pt(8.5)
    run.font.color.rgb = GREY


def add_code_box(doc, lines):
    """Add a monospaced, shaded box used for tool-output excerpts."""
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    shade(cell, CODE_FILL)
    cell.text = ""
    first = True
    for line in lines:
        paragraph = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0
        run = paragraph.add_run(line if line else " ")
        run.font.name = "Consolas"
        run.font.size = Pt(8.0)
    doc.add_paragraph()


def add_paragraph(doc, text, size=10.5, bold=False, italic=False,
                  space_after=6, color=None):
    """Add a normal paragraph with the requested formatting."""
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(space_after)
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    return paragraph


def add_bullets(doc, items):
    """Add a bullet list of plain strings."""
    for item in items:
        paragraph = doc.add_paragraph(style="List Bullet")
        paragraph.paragraph_format.space_after = Pt(2)
        run = paragraph.add_run(item)
        run.font.size = Pt(10)


def page_field(paragraph):
    """Append a PAGE field to a footer paragraph."""
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)
    run.font.size = Pt(8.5)
    run.font.color.rgb = GREY


# --------------------------------------------------------------------------
# report content
# --------------------------------------------------------------------------
def main():  # noqa: C901 - a linear build script
    """Assemble and save Lab5_Static_Code_Analysis_Report.docx."""
    before = {
        "pylint": REPORTS / "before" / "pylint_report.txt",
        "bandit": REPORTS / "before" / "bandit_report.txt",
        "flake8": REPORTS / "before" / "flake8_report.txt",
    }
    stage2 = REPORTS / "stages" / "after_pylint_fix"
    stage3 = REPORTS / "stages" / "after_bandit_fix"
    after = {
        "pylint": REPORTS / "after" / "pylint_report.txt",
        "bandit": REPORTS / "after" / "bandit_report.txt",
        "flake8": REPORTS / "after" / "flake8_report.txt",
    }

    pylint_before = pylint_messages(before["pylint"])
    bandit_before = bandit_issues(before["bandit"])
    flake8_before = flake8_messages(before["flake8"])
    score_before = pylint_score(before["pylint"])
    score_after = pylint_score(after["pylint"])
    severity_of = {code: severity for _, code, severity in bandit_before}

    assert len(pylint_before) == 34, len(pylint_before)
    assert len(bandit_before) == 9, len(bandit_before)
    assert len(flake8_before) == 17, len(flake8_before)
    assert score_before == "5.64" and score_after == "10.00"

    # ---------- issue table rows ----------
    issue_rows = [
        ("1", "Pylint C0114 / C0116\nMissing docstrings (Low)", "C0114 C0116",
         "Module and 11 functions have no docstrings; intent is undocumented.",
         "Added a module docstring and a one-line docstring for every function."),
        ("2", "Pylint W0102\nMutable default argument (High)", "W0102",
         "logs=[] is created once and shared, so entries leak between calls.",
         "Default changed to None; a fresh list is created inside the function."),
        ("3", "Pylint W0702 / Flake8 E722\nBare except (Medium)", "W0702 E722",
         "except: catches SystemExit and KeyboardInterrupt and hides errors.",
         "Catch only OSError, ValueError and SyntaxError where expected."),
        ("4", "Pylint W0718\nBroad exception caught (Medium)", "W0718",
         "except Exception hides programming errors and makes debugging hard.",
         "Narrowed to OSError; the offending file name is included in the message."),
        ("5", "Pylint W0602 / Flake8 F824\nRedundant global (Medium)", "W0602 F824",
         "global users is declared although the name is never assigned in scope.",
         "Removed the statement; mutating the dictionary does not require global."),
        ("6", "Pylint W0611 / Flake8 F401\nUnused imports (Low)", "W0611 F401",
         "json and shutil are imported but never referenced.",
         "Removed both imports; shutil returns later for the safe copy fix."),
        ("7", "Pylint W0612 / Flake8 F841\nUnused variables (Medium)", "W0612 F841",
         "name, email, unused_value and i are assigned but never read.",
         "Deleted the dead assignments and used _ for the throw-away loop counter."),
        ("8", "Pylint C0121 / Flake8 E711, E712\nSingleton comparisons (Low)",
         "C0121 E711 E712",
         "== None, == True and == False compare by equality instead of identity.",
         "Used is None, the truthiness of login() and not os.path.exists()."),
        ("9", "Pylint R1703 / R1705\nSimplifiable if / else (Low)", "R1703 R1705",
         "if/else returns literal booleans and repeats the condition inverted.",
         "Returned the comparison expression directly."),
        ("10", "Pylint R1710\nInconsistent return statements (Medium)", "R1710",
         "read_config returns a value on success but implicitly None on failure.",
         "Added an explicit return None in the exception branch."),
        ("11", "Pylint R1732 / W1514\nFile handling (Medium)", "R1732 W1514",
         "open() is used without a context manager and without an encoding.",
         "Used with open(path, encoding=\"utf-8\") in both functions."),
        ("12", "Pylint C0410 / Flake8 E401\nMultiple imports (Low)", "C0410 E401",
         "os and sys are imported on a single line.",
         "Split the import into one statement per module."),
        ("13", "Pylint W0123 / Bandit B307\neval() (High)", "W0123 B307",
         "eval() executes arbitrary code read from config.txt (remote code execution).",
         "Replaced with ast.literal_eval, which parses literals only."),
        ("14", "Bandit B324\nWeak MD5 hash (High)", "B324",
         "MD5 is broken; unsalted digests are fast to crack and easy to collide.",
         "Salted SHA-256; production code should use PBKDF2, bcrypt or argon2."),
        ("15", "Bandit B605\nShell injection via os.system (High)", "B605",
         "File names are concatenated into a shell string; metacharacters execute.",
         "Replaced os.system with shutil.copy2(src, dst)."),
        ("16", "Bandit B602\nsubprocess with shell=True (High)", "B602",
         "An arbitrary command string is passed to the Windows shell.",
         "Removed the unused deploy helper; no shell execution remains."),
        ("17", "Bandit B105\nHardcoded password (Low)", "B105",
         "The admin password admin123 is committed inside the source file.",
         "Removed the unused constant; secrets belong in environment variables."),
        ("18", "Bandit B311\nInsecure random token (Low)", "B311",
         "random.random() is predictable and unsuitable for session tokens.",
         "secrets.token_hex(16) provides 128-bit cryptographically secure tokens."),
        ("19", "Bandit B101\nassert for validation (Low)", "B101",
         "Assertions are stripped when Python runs with -O, losing the check.",
         "Replaced with if len(password) < 8: raise ValueError(...)."),
        ("20", "Bandit B108\nHardcoded /tmp directory (Medium)", "B108",
         "A fixed /tmp path is predictable and open to symlink attacks.",
         "tempfile.gettempdir() joined with lab5 gives a portable safe location."),
        ("21", "Bandit B404\nsubprocess import (Low)", "B404",
         "The subprocess module is imported only for the unsafe deploy helper.",
         "Removed together with the unsafe command-running function."),
        ("22", "Flake8 E501\nLine too long (Low)", "E501",
         "The error message line is 87 characters, above the PEP 8 limit of 79.",
         "Stored the message in a variable and printed it in two lines."),
        ("23", "Flake8 E302, E303, E231, E225, W291 / Pylint C0303\n"
               "PEP 8 layout (Low)", "C0303 E302 E303 E231 E225 W291",
         "Wrong blank-line counts, missing spaces around , and =, trailing spaces.",
         "Reformatted the file to PEP 8 (blank lines, spacing, no trailing spaces)."),
    ]

    table_rows = []
    for number, title, codes, desc, fix in issue_rows:
        wanted = codes.split()
        lines = codes_to_lines(
            pylint_before + flake8_before
            + [(line, code) for line, code, _ in bandit_before],
            wanted,
        )
        # keep only codes that actually exist in the before reports
        table_rows.append((number, title, fmt_lines(lines), desc, fix))

    # ---------- document ----------
    doc = Document()
    doc.core_properties.title = "Lab 5 - Static Code Analysis"
    doc.core_properties.subject = "Python code quality with Pylint, Bandit and Flake8"
    doc.core_properties.author = "Lab 5 Student"

    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.7)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    styles = doc.styles
    styles["Normal"].font.name = "Calibri"
    styles["Normal"].font.size = Pt(10.5)
    for style_name, size in (("Heading 1", 15), ("Heading 2", 12)):
        style = styles[style_name]
        style.font.name = "Calibri"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = DARK_BLUE
        style.paragraph_format.space_before = Pt(10)
        style.paragraph_format.space_after = Pt(4)

    # header and footer
    header_paragraph = section.header.paragraphs[0]
    header_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header_run = header_paragraph.add_run(
        "Lab 5  •  Static Code Analysis  •  Python 3.12"
    )
    header_run.font.size = Pt(8)
    header_run.font.color.rgb = GREY
    border = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "BFBFBF")
    border.append(bottom)
    header_paragraph._p.get_or_add_pPr().append(border)

    footer_paragraph = section.footer.paragraphs[0]
    footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer_run = footer_paragraph.add_run(
        "Lab 5 — Static Code Analysis  |  Page "
    )
    footer_run.font.size = Pt(8.5)
    footer_run.font.color.rgb = GREY
    page_field(footer_paragraph)

    # ---------------- cover ----------------
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(90)
    add_paragraph(doc, "LAB 5", size=13, bold=True,
                  color=ACCENT, space_after=2)
    add_paragraph(doc, "Static Code Analysis", size=28, bold=True,
                  color=DARK_BLUE, space_after=2)
    add_paragraph(doc, "Improving Python quality, security and style with "
                       "Pylint, Bandit and Flake8",
                  size=12, italic=True, color=GREY, space_after=24)

    cover_rows = [
        ("Course", "Software Engineering / Secure Coding"),
        ("Student name", "________________________________"),
        ("Roll / Reg. no.", "________________________________"),
        ("Date", "27 October 2026"),
        ("Repository", "Lab5_Static_Code_Analysis (Git)"),
    ]
    cover_table = doc.add_table(rows=0, cols=2)
    cover_table.style = "Table Grid"
    cover_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cover_table.autofit = False
    for label, value in cover_rows:
        row = cover_table.add_row()
        no_split(row)
        cell_text(row.cells[0], label, bold=True, size=9.5, color=DARK_BLUE)
        cell_text(row.cells[1], value, size=9.5)
        shade(row.cells[0], ZEBRA_FILL)
        row.cells[0].width = Cm(4.2)
        row.cells[1].width = Cm(12.4)
    cover_table.columns[0].width = Cm(4.2)
    cover_table.columns[1].width = Cm(12.4)

    add_paragraph(doc, "", space_after=10)
    add_paragraph(
        doc,
        "This report documents a baseline Python utility, the issues reported "
        "against it by three industry-standard static analysis tools, the fixes "
        "applied, and the verification that the final code passes all three "
        "checks: Pylint above 9.5, zero Bandit findings and zero Flake8 findings.",
        size=10, italic=True, color=GREY,
    )

    # ---------------- 1. objective ----------------
    doc.add_page_break()
    doc.add_heading("1. Objective", level=1)
    add_paragraph(doc, (
        "The objective of this lab is to enhance Python code quality, security "
        "and style by using static analysis tools — Pylint, Bandit and Flake8 — "
        "to detect and rectify common programming issues in an assigned Python "
        "script. Static analysis examines source code without executing it, so "
        "it can expose defect and vulnerability patterns that a quick manual "
        "review often misses."
    ))
    add_paragraph(doc, "The lab requires the following deliverables:")
    add_bullets(doc, [
        "A cleaned and updated version of the script with at least four issues fixed.",
        "A filled-out table documenting the identified issues and how they were addressed.",
        "Answers to the reflection questions, based on the hands-on experience.",
        "Pylint, Bandit and Flake8 reports captured before and after the fixes.",
        "Three commits — one for the fixes raised by each of the three tools.",
    ])

    # ---------------- 2. environment ----------------
    doc.add_heading("2. Environment and Tools", level=1)
    add_paragraph(doc, (
        "The analysis was performed with Python 3.12.14 on Windows 11. The three "
        "tools were installed with pip; they are complementary rather than "
        "redundant, as the table below shows."
    ))
    add_table(
        doc,
        ["Tool", "Version", "Command used", "Role in the analysis"],
        [
            ("Pylint", "4.1.3",
             "pylint user_utils.py",
             "Strict reviewer: correctness, design and documentation, with a weighted 10-point score."),
            ("Bandit", "1.9.4",
             "bandit -r user_utils.py",
             "Security guard: dangerous calls, weak crypto, injection and hardcoded secrets."),
            ("Flake8", "7.4.1",
             "flake8 user_utils.py",
             "PEP 8 grammar checker: formatting, whitespace, line length, unused names."),
        ],
        [2.6, 2.0, 6.4, 6.6],
        font_size=8.5,
    )
    caption(doc, "Table 1: Static analysis tools and the aspect of code quality each one targets.")
    add_paragraph(doc, (
        "The environment setup follows the handout, adapted to Windows "
        "PowerShell (the handout shows the POSIX activation command):"
    ))
    add_code_box(doc, [
        "python -m venv venv",
        "venv\\Scripts\\activate          # source venv/bin/activate on Linux/macOS",
        "pip install pylint bandit flake8",
        "python user_utils.py",
    ])
    caption(doc, "Listing 1: Environment setup and the command used to run the baseline program.")

    # ---------------- 3. program ----------------
    doc.add_heading("3. Program Under Analysis", level=1)
    add_paragraph(doc, (
        "The baseline script, original_script.py (106 lines), is a small "
        "user-management utility. It keeps an in-memory user registry, hashes "
        "passwords, issues session tokens, reads a configuration file, copies "
        "files, runs deployment commands, validates passwords and parses a "
        "comma-separated user file. A main() function demonstrates the helpers "
        "using sys.argv defaults, so the program runs end to end from the "
        "terminal."
    ))
    add_paragraph(doc, (
        "Note: the GitRepo-linked script referenced by the handout was not part "
        "of the provided materials. The baseline was therefore written to "
        "reproduce every issue class named in the lab — mutable default "
        "arguments, overly broad exceptions and insecure coding patterns — so "
        "that the analysis, the documented table and the fixes are fully "
        "reproducible from the submitted files."
    ), size=9.5, italic=True, color=GREY)
    add_paragraph(doc, "The baseline deliberately contains four groups of defects:")
    add_bullets(doc, [
        "Correctness defects: a mutable default argument that leaks log entries between calls, and an inconsistent return in read_config.",
        "Exception-handling defects: bare except and a broad except Exception that hide the real cause of failures.",
        "Security defects: eval() on file content, an unsalted MD5 password hash, os.system() command injection, subprocess with shell=True, a hardcoded admin password, predictable session tokens, assert-based validation and a hardcoded /tmp directory.",
        "Style defects: missing docstrings and type-obvious names, unused imports and variables, singleton comparisons, wrong blank-line counts and an 87-character line.",
    ])

    # ---------------- 4. before ----------------
    doc.add_heading("4. Static Analysis Results — Before Fixes", level=1)
    add_paragraph(doc, (
        "Each tool was run against the unmodified script and its output was "
        "redirected to a text file, exactly as the handout describes. The raw "
        "reports are included in reports/before/."
    ))
    add_table(
        doc,
        ["Tool", "Raw findings", "Score / severity", "Interpretation"],
        [
            ("Pylint", "34 messages",
             "5.64 / 10",
             "C: 17, W: 12, R: 5 — well below the required 9.5."),
            ("Bandit", "9 issues",
             "3 High, 2 Medium, 4 Low",
             "3 high-severity findings, all exploitable in principle."),
            ("Flake8", "17 violations",
             "9 errors, 6 pyflakes, 1 warning",
             "PEP 8 layout plus dead code and undefined behaviour risks."),
        ],
        [2.4, 3.6, 4.6, 6.4],
        font_size=8.5,
    )
    caption(doc, "Table 2: Initial findings reported by the three tools against original_script.py.")
    add_paragraph(doc, (
        "Pylint's score is a weighted average of message categories; the "
        "high-severity Bandit findings (weak MD5 hash, os.system injection and "
        "shell=True) were chosen as the first priority, followed by the mutable "
        "default argument, which is a genuine runtime bug."
    ))

    # ---------------- 5. documented issues ----------------
    doc.add_heading("5. Documented Issues and Fixes", level=1)
    add_paragraph(doc, (
        "The following table documents the issues found in the baseline script "
        "and the fix applied to each one. Line numbers refer to "
        "original_script.py. For Pylint and Flake8 findings the severity is our "
        "impact assessment (High = bug or exploit, Medium = maintainability or "
        "hidden failure, Low = style or documentation); for Bandit findings the "
        "severity is the one reported by the tool."
    ))
    add_table(
        doc,
        ["#", "Tool / rule — issue (severity)", "Line(s)",
         "Description", "Fix applied"],
        table_rows,
        [0.8, 4.0, 2.1, 4.7, 4.6],
        font_size=7.6,
    )
    caption(doc, (
        "Table 3: Findings documented and fixed (23 grouped rows covering all 60 "
        "raw findings: 34 Pylint messages, 9 Bandit issues and 17 Flake8 violations)."
    ))

    # ---------------- 6. fixes ----------------
    doc.add_heading("6. Fixes Applied", level=1)
    add_paragraph(doc, (
        "The fixes were made in three commits, one per tool, following the order "
        "of the lab: first the Pylint findings, then the Bandit security "
        "findings, then the Flake8 formatting findings. Working tool by tool "
        "kept each review focused and made it easy to verify that a fix for one "
        "tool had not introduced a new finding in another."
    ))
    add_table(
        doc,
        ["Commit", "Tool addressed", "Summary of changes", "Files changed"],
        [
            ("424a8b0", "— (baseline)",
             "Committed the buggy script, sample data and the three \"before\" reports.",
             "7 files"),
            ("fa5586d", "Pylint",
             "Docstrings, mutable default argument, exception types, global statement, "
             "unused imports/variables, comparisons, file context managers.",
             "4 files"),
            ("b2b2b31", "Bandit",
             "Salted SHA-256, secrets tokens, ast.literal_eval, shutil.copy2, "
             "explicit ValueError, tempfile, removed hardcoded secret and subprocess.",
             "4 files"),
            ("c3ecc1a", "Flake8",
             "PEP 8 blank lines, spacing around comma/operator, split the long line, "
             "regenerated the final reports.",
             "4 files"),
        ],
        [2.2, 3.0, 8.2, 3.6],
        font_size=8.0,
    )
    caption(doc, "Table 4: Git history — three fix commits, one per analysis tool, after the baseline commit.")

    add_paragraph(doc, "Key design decisions during the fixes:", bold=True)
    add_bullets(doc, [
        "Root-cause fixes instead of suppression: eval() was replaced with ast.literal_eval and shell command execution was removed, rather than silencing the rules with # nosec or # pylint: disable.",
        "Standard library replacements: os.system(\"cp …\") became shutil.copy2(), which is portable, quoting-safe and has no shell.",
        "Defence in depth for passwords: the MD5 digest became a salted SHA-256 digest; the report notes that PBKDF2, bcrypt or argon2 are the production choice for password storage.",
        "Dead code removed: the unused admin-password constant, the unused deploy() helper, unused imports and dead local variables were deleted rather than given dummy uses — deleting dead code is safer than keeping it.",
        "Explicit failure modes: a short password now raises ValueError instead of relying on assert, and read_config returns None explicitly on failure instead of falling off the end of the function.",
    ])

    # ---------------- 7. verification ----------------
    doc.add_heading("7. Verification — After Fixes", level=1)
    add_paragraph(doc, (
        "All three tools were re-run after every commit. The final state below "
        "meets every target from Step 4 of the handout: Pylint above 9.5, all "
        "Bandit findings handled, and all Flake8 violations fixed. The final "
        "reports are stored in reports/after/."
    ))
    add_table(
        doc,
        ["Tool", "Before", "After", "Requirement"],
        [
            ("Pylint", "34 messages, 5.64 / 10", "0 messages, 10.00 / 10",
             "Score > 9.5 — PASS"),
            ("Bandit", "9 issues (3 High, 2 Medium, 4 Low)", "0 issues",
             "All issues handled — PASS"),
            ("Flake8", "17 violations", "0 violations",
             "All violations fixed — PASS"),
        ],
        [2.4, 5.2, 4.6, 4.8],
        font_size=8.5,
    )
    caption(doc, "Table 5: Before/after comparison of the three tools; every handout target is met.")

    add_paragraph(doc, "The intermediate verification after each commit confirms that each step improved only its own tool's findings:", bold=True)
    add_table(
        doc,
        ["Stage", "Pylint", "Bandit", "Flake8", "Tool fixed"],
        [
            ("Baseline (424a8b0)", "5.64/10 (34)", "9 issues", "17 violations", "—"),
            ("After fa5586d", "10.00/10 (0)", "8 issues", "5 violations", "Pylint"),
            ("After b2b2b31", "10.00/10 (0)", "0 issues", "5 violations", "Bandit"),
            ("After c3ecc1a", "10.00/10 (0)", "0 issues", "0 violations", "Flake8"),
        ],
        [4.4, 3.4, 2.8, 3.0, 3.0],
        font_size=8.5,
    )
    caption(doc, "Table 6: Findings count after each commit; the final row is the submitted state.")

    add_paragraph(doc, "Final tool output (excerpts):", bold=True)
    add_code_box(doc, [
        "$ pylint user_utils.py",
        "Your code has been rated at 10.00/10 (previous run: 5.64/10, +4.36)",
        "",
        "$ bandit -r user_utils.py",
        "Test results:",
        "        No issues identified.",
        "        Total issues (by severity): Low 0, Medium 0, High 0",
        "",
        "$ flake8 user_utils.py",
        "(no output — exit code 0)",
    ])
    caption(doc, "Listing 2: Verification output for the final version of user_utils.py.")

    # ---------------- 8. reflection ----------------
    doc.add_heading("8. Reflection", level=1)
    add_paragraph(doc, (
        "The reflection questions below follow the learning outcomes printed in "
        "the handout; they are answered from the actual run of the three tools "
        "on this script."
    ))

    reflections = [
        ("Q1. What is static code analysis, and how does it differ from testing?",
         "Static analysis inspects the source code, bytecode or configuration "
         "without running the program. It recognises patterns that are known to "
         "be risky — a mutable default argument, an except clause that is too "
         "broad, MD5 used for passwords, a shell command built by string "
         "concatenation — and reports them deterministically, even on code paths "
         "that a test suite never executes. Testing, by contrast, executes the "
         "program with particular inputs and checks observable behaviour. The "
         "two are complementary: static analysis cannot tell whether a function "
         "returns the right value for a given input, and tests rarely prove the "
         "absence of a dangerous construct. In this lab static analysis found "
         "nine security defects and a real runtime bug that a functional test of "
         "main() would not have exposed, because the buggy functions were barely "
         "exercised by the happy path."),
        ("Q2. Which of the three tools produced the most valuable findings?",
         "Bandit produced the most valuable findings in terms of risk: three "
         "high-severity issues (weak MD5 password hashing, command injection via "
         "os.system, and shell=True in subprocess) are directly exploitable and "
         "were invisible to Pylint and Flake8. Pylint produced the most valuable "
         "findings for correctness: the mutable default argument is a genuine "
         "bug that would corrupt the audit log between calls, and the broad "
         "except clauses would hide failures in production. Flake8's contribution "
         "was consistency and readability rather than risk reduction. The "
         "lesson is that the tools are complementary: security, correctness and "
         "style are different questions, and a project should run all three."),
        ("Q3. How were the findings prioritised and sequenced?",
         "Severity was the first sorting key. The three high-severity Bandit "
         "findings and the mutable-default bug were treated as release blockers; "
         "medium findings (broad exceptions, file handles, insecure temp path) "
         "came next; style and documentation came last. Sequencing followed the "
         "tool boundaries of the lab: the Pylint commit touched logic and "
         "documentation, the Bandit commit replaced unsafe constructs with safe "
         "standard-library equivalents, and the Flake8 commit normalised layout. "
         "After every commit all three tools were re-run, because removing a "
         "dangerous call can change line numbers and reveal new findings — and "
         "because the Pylint score can drop if a refactor introduces an unused "
         "variable or a missing docstring."),
        ("Q4. What challenges were encountered, and how were they resolved?",
         "The hardest part was fixing causes rather than symptoms. It is easy to "
         "make Bandit quiet with # nosec or Pylint quiet with a disable comment, "
         "but that hides the risk; instead eval() was replaced with "
         "ast.literal_eval, the shell-based backup was replaced with "
         "shutil.copy2, and the unused command runner was deleted. A second "
         "challenge was cross-platform behaviour: the handout's cp command does "
         "not exist on Windows, which is itself evidence of why os.system was a "
         "bad design. A third challenge was the different thresholds of the "
         "tools — Pylint accepts lines up to 100 characters while Flake8 enforces "
         "79, so a file can be Pylint-clean and still fail Flake8. That "
         "discrepancy showed why a project should pin tool configuration in "
         "setup.cfg, pyproject.toml or tox.ini so that local results and CI "
         "results agree."),
        ("Q5. What are the limitations of static analysis?",
         "Static analysis reasons about patterns, not about intent or runtime "
         "state. It reports false positives (for example, importing subprocess "
         "is not automatically a vulnerability), it cannot prove business-logic "
         "correctness, and it cannot detect problems that require data-flow "
         "across services or concurrency reasoning. It also cannot replace "
         "review: a human still has to decide whether a finding matters and "
         "whether the fix preserves behaviour. The findings also depend on the "
         "tool version and rule configuration, so a clean report only means "
         "\"clean for these rules\" — which is why the reports in this lab "
         "record the exact commands and versions used."),
        ("Q6. How did static analysis improve this script?",
         "The improvements are measurable: the Pylint score rose from 5.64 to "
         "10.00, Bandit findings dropped from nine to zero, and Flake8 "
         "violations dropped from seventeen to zero. Beyond the numbers, the "
         "code became safer (no dynamic code execution, no shell, a strong "
         "salted hash, secure tokens), more portable (tempfile and shutil "
         "instead of /tmp and cp), more explicit (explicit ValueError and "
         "return None instead of assert and implicit fall-through) and easier to "
         "maintain (docstrings on every function, context managers, PEP 8 "
         "layout). The exercise also demonstrated how static analysis fits into "
         "a workflow: run it before review, fix high-severity items first, and "
         "gate merges on the same commands so that regressions are caught "
         "automatically."),
    ]
    for question, answer in reflections:
        doc.add_heading(question, level=2)
        add_paragraph(doc, answer)

    # ---------------- 9. conclusion ----------------
    doc.add_heading("9. Conclusion", level=1)
    add_paragraph(doc, (
        "This lab turned a deliberately flawed utility into a clean, safe and "
        "documented script. The baseline produced 60 findings across the three "
        "tools: 34 Pylint messages (score 5.64/10), nine Bandit security issues "
        "and seventeen Flake8 violations. After three focused fix commits, the "
        "final script scores 10.00/10 in Pylint, reports no Bandit issues and "
        "passes Flake8 without a single warning, exceeding every target set in "
        "the handout. More importantly, the exercise showed what each tool is "
        "good at — Pylint for correctness and design, Bandit for security and "
        "Flake8 for consistency — and why running all three, before review and "
        "in CI, catches classes of defects that manual reading and functional "
        "testing routinely miss."
    ))

    # ---------------- appendix A ----------------
    doc.add_page_break()
    doc.add_heading("Appendix A. Submitted Files", level=1)
    add_paragraph(doc, (
        "All files referenced in this report are in the submission directory "
        "Lab5_Static_Code_Analysis/."
    ))
    add_table(
        doc,
        ["Path", "Contents", "Deliverable"],
        [
            ("original_script.py",
             "The buggy 106-line baseline analysed in this report.", "Baseline"),
            ("user_utils.py",
             "The cleaned and updated script; passes all three tools.", "Deliverable 1"),
            ("reports/before/*.txt",
             "pylint_report.txt, bandit_report.txt, flake8_report.txt before fixes.",
             "Deliverable 2"),
            ("reports/after/*.txt",
             "The same three reports after fixes (empty Flake8 output = clean).",
             "Deliverable 2"),
            ("reports/stages/",
             "Intermediate reports after the Pylint fix and after the Bandit fix.",
             "Evidence"),
            ("config.txt, users.txt",
             "Sample data used to run the program end to end.", "Support"),
            ("Lab5_Static_Code_Analysis_Report.docx",
             "This document: issue table, fixes, verification and reflection.",
             "Deliverable 3"),
            ("build_report.py",
             "Script that regenerates this DOCX from the reports and fixed data.",
             "Support"),
            ("README.md",
             "One-page guide to the repository and the exact commands used.",
             "Support"),
        ],
        [6.0, 7.8, 3.2],
        font_size=8.5,
    )
    caption(doc, "Table 7: Inventory of the submitted files and the deliverable each one satisfies.")

    # ---------------- appendix B ----------------
    doc.add_heading("Appendix B. Reproducing the Results", level=1)
    add_paragraph(doc, "From the Lab5_Static_Code_Analysis directory:")
    add_code_box(doc, [
        "# 1. environment",
        "python -m venv venv",
        "venv\\Scripts\\activate",
        "pip install pylint bandit flake8",
        "",
        "# 2. before the fixes (baseline snapshot)",
        "pylint original_script.py > reports\\before\\pylint_report.txt",
        "bandit -r original_script.py > reports\\before\\bandit_report.txt",
        "flake8 original_script.py > reports\\before\\flake8_report.txt",
        "",
        "# 3. after the fixes",
        "pylint user_utils.py > reports\\after\\pylint_report.txt",
        "bandit -r user_utils.py > reports\\after\\bandit_report.txt",
        "flake8 user_utils.py > reports\\after\\flake8_report.txt",
        "",
        "# 4. inspect history",
        "git log --oneline",
        "",
        "# 5. rebuild this report",
        "python build_report.py",
    ])
    caption(doc, "Listing 3: Commands used to produce every artefact in this submission.")

    add_paragraph(doc, (
        "Note: on Windows PowerShell, redirection writes UTF-16 by default. The "
        "reports committed to the repository were saved as UTF-8 text; the "
        "commands above work as shown for a normal terminal, but PowerShell "
        "users may prefer Out-File -Encoding utf8 to keep the reports readable "
        "in Git and in editors."
    ), size=9.0, italic=True, color=GREY)

    # ---------------- appendix C ----------------
    doc.add_page_break()
    doc.add_heading("Appendix C. Deliverables Checklist", level=1)
    add_paragraph(doc, "Table 8 maps every requirement in the handout to the "
                       "evidence that satisfies it.", space_after=4)
    add_table(
        doc,
        ["Requirement", "Where it is satisfied", "Status"],
        [
            ("Cleaned script, at least four issues fixed",
             "user_utils.py — all 60 raw findings addressed (Table 3).", "Complete"),
            ("Filled-out issue table",
             "Section 5, Table 3 (23 rows).", "Complete"),
            ("Answers to reflection questions",
             "Section 8, Q1–Q6.", "Complete"),
            ("Pylint, Bandit, Flake8 reports (before and after)",
             "reports/before/ and reports/after/.", "Complete"),
            ("Pylint score above 9.5",
             "10.00/10 in reports/after/pylint_report.txt.", "Complete"),
            ("All Bandit issues handled",
             "\"No issues identified.\" in reports/after/bandit_report.txt.", "Complete"),
            ("All Flake8 issues fixed",
             "Empty output (exit code 0) in reports/after/flake8_report.txt.", "Complete"),
            ("Three commits, one per tool",
             "fa5586d (Pylint), b2b2b31 (Bandit), c3ecc1a (Flake8).", "Complete"),
        ],
        [7.0, 7.6, 2.4],
        font_size=8.5,
    )

    out = LAB / "Lab5_Static_Code_Analysis_Report.docx"
    doc.save(out)
    print(f"wrote {out} ({out.stat().st_size} bytes)")
    print(f"pylint before={len(pylint_before)} bandit before={len(bandit_before)} "
          f"flake8 before={len(flake8_before)} rows={len(table_rows)}")


if __name__ == "__main__":
    main()
