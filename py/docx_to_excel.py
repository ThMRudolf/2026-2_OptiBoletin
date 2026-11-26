#!/usr/bin/env python3
"""
docx_to_excel.py
════════════════
Reads every .doc/.docx in a folder, parses course listings with pure Python
(python-docx), deduplicates by course code + title, and writes a formatted
Excel with columns:
    A Departamento · B Código · C Título · D Profesor · E Prerrequisitos · F Descripción
Missing fields are left blank. "Notas al plan de estudios" files are skipped.
Legacy .doc files are converted via MS Word (pywin32) or LibreOffice.

No external API required.

Dependencies
────────────
    pip install python-docx openpyxl
    pip install pywin32          # optional, only for .doc files

Usage
─────
    python docx_to_excel.py                           # folder picker → Optativas_<Semestre>_<Año>.xlsx
    python docx_to_excel.py -s Primavera -a 2027      # → Optativas_Primavera_2027.xlsx
    python docx_to_excel.py -i ./docx_folder
    python docx_to_excel.py -i ./docs -o Optativas_Primavera_2027.xlsx
    python docx_to_excel.py -i ./docs -v              # verbose
"""

import re, sys, argparse
from pathlib import Path

try:
    from docx import Document
except ImportError:
    sys.exit("Missing dependency:  pip install python-docx")

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Missing dependency:  pip install openpyxl")


# ─── DEPARTMENT LOOKUP ────────────────────────────────────────────────────────

DEPT_ORDER = [
    "Administración", "Bienestar / Formación General", "Ciencia Política",
    "Computación", "Contabilidad", "Derecho", "Economía", "Estadística",
    "Estudios Generales", "Estudios Internacionales",
    "Ingeniería Eléctrica y Electrónica",
    "Ingeniería Industrial y Operaciones", "Matemáticas",
]

PREFIX_TO_DEPT = {
    "ADM": "Administración",    "CEB": "Bienestar / Formación General",
    "CSO": "Ciencia Política",  "COM": "Computación",
    "CON": "Contabilidad",      "DER": "Derecho",
    "ECO": "Economía",          "EST": "Estadística",
    "EGN": "Estudios Generales","EIN": "Estudios Internacionales",
    "SDI": "Ingeniería Eléctrica y Electrónica",
    "IEE": "Ingeniería Eléctrica y Electrónica",
    "IIO": "Ingeniería Industrial y Operaciones",
    "MAT": "Matemáticas",
}
COURSE_PREFIXES = set(PREFIX_TO_DEPT.keys())

def dept_from_code(code: str) -> str:
    return PREFIX_TO_DEPT.get(code.split("-")[0].upper(), code.split("-")[0])


# ─── REGEX PATTERNS ──────────────────────────────────────────────────────────

# Course code at start — digits may be immediately followed by title text
# (no separator) because iter_lines collapses whitespace.
_CODE_AT_START = re.compile(
    r'^\s*(?:(?:CLAVE|MATERIA)\s*[:\.]?\s*)?'   # optional CLAVE: / MATERIA:
    r'([A-Z]{2,4})'                  # prefix
    r'[\s\-–—\u00a0]*'              # optional separator
    r'(\d{4,6})'                     # 4-6 digits
    r'(?:[A-Z](?![A-Za-zÀ-ÿ]))?',    # optional 1-letter suffix (not a title's 1st letter)
    re.UNICODE | re.IGNORECASE,
)

PROF_RE   = re.compile(r'^\s*PROFESOR(?:ES|AS|A)?\s*[:\.]?\s*',  re.IGNORECASE)
PREREQ_RE = re.compile(r'^\s*(?:SIN\s+)?PRERREQUISITOS?\s*[:\.]?\s*', re.IGNORECASE)
DESC_RE   = re.compile(r'^\s*DESCRIPCI[ÓO]N\s*[:\.]?\s*',  re.IGNORECASE)
NOMBRE_RE = re.compile(r'^\s*NOMBRE\s+DEL\s+CURSO\s*[:\.]?\s*', re.IGNORECASE)
CLAVE_RE  = re.compile(r'^\s*CLAVE\s*[:\.]?\s*',            re.IGNORECASE)

SKIP_RE   = re.compile(
    r'^\s*(?:departamento|académico|materias|optativas|otoño|primavera|verano'
    r'|invierno|estimados|directores|a continuación|fecha|semestre|oferta'
    r'|licenciatura|optativas oto|temario|\d+[\.\)])',
    re.IGNORECASE | re.UNICODE,
)


# ─── PARAGRAPH ITERATOR ──────────────────────────────────────────────────────

def iter_lines(doc) -> list[str]:
    """
    Yield clean text lines from body paragraphs and table cells, in order.
    Tabs are replaced with a SPACE (not stripped) so CODE\tTITLE stays parseable.
    Embedded newlines are split into separate lines.
    """
    from docx.oxml.ns import qn

    def para_text(elem):
        return "".join(r.text or "" for r in elem.iter(qn("w:t")))

    raw_lines = []
    body = doc.element.body
    for child in body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "p":
            raw_lines.append(para_text(child))
        elif tag == "tbl":
            for tr in child.iter(qn("w:tr")):
                seen: set[str] = set()
                for tc in tr.iter(qn("w:tc")):
                    ct = " ".join(
                        para_text(p) for p in tc.iter(qn("w:p"))
                    ).strip()
                    key = ct[:50]
                    if ct and key not in seen:
                        seen.add(key)
                        raw_lines.append(ct)

    result = []
    for line in raw_lines:
        # Replace non-breaking space and tabs with a regular space
        line = line.replace("\xa0", " ").replace("\t", " ")
        # Split on embedded newlines
        for sub in line.split("\n"):
            # Collapse multiple spaces to one
            s = re.sub(r" {2,}", " ", sub).strip()
            if s:
                result.append(s)
    return result


# ─── CODE EXTRACTION ─────────────────────────────────────────────────────────

def extract_code_at_start(text: str):
    """
    Return (code, title_remainder) if line starts with a valid course code.
    title_remainder is everything after the code (may be empty).
    Returns (None, None) if not a course-code line.
    """
    m = _CODE_AT_START.match(text)
    if not m:
        return None, None
    prefix = m.group(1).upper()
    if prefix not in COURSE_PREFIXES:
        return None, None
    code = f"{prefix}-{m.group(2)}"
    # remainder starts right after the full match
    remainder = text[m.end():].strip().lstrip("-–—:. ")
    return code, remainder


# ─── PARSER ──────────────────────────────────────────────────────────────────

def parse_docx(path: Path) -> list[dict]:
    doc   = Document(str(path))
    lines = iter_lines(doc)

    courses: list[dict] = []
    cur: dict | None    = None
    desc_parts: list[str] = []
    state = None  # "prof" | "prereq" | "desc" | None

    def flush():
        nonlocal cur, desc_parts
        if cur and cur.get("code"):
            if desc_parts:
                d = " ".join(desc_parts).strip()
                cur["description"] = re.sub(r'^Breve descripción del contenido\.?\s*', '',
                                            d, flags=re.IGNORECASE)
            cur["professor"]     = cur["professor"].strip().rstrip(".,;/ ")
            cur["prerequisites"] = cur["prerequisites"].strip()
            courses.append(cur)
        cur = None
        desc_parts = []

    def new_course(code, title="", prof=""):
        return {"code": code, "title": title, "professor": prof,
                "prerequisites": "", "description": "",
                "department": dept_from_code(code)}

    def label_ahead(i, n=3):
        """True if a PROFESOR label follows within n lines, before any
        DESCRIPCIÓN label or other course-code line."""
        for l in lines[i + 1:i + 1 + n]:
            if PROF_RE.match(l):
                return True
            if DESC_RE.match(l) or extract_code_at_start(l)[0]:
                return False
        return False

    for i, line in enumerate(lines):
        # ── Section headings (ALL CAPS line directly before a course code) ──
        nxt = lines[i + 1] if i + 1 < len(lines) else ""
        if (len(line) < 60 and not re.search(r"[a-zà-ÿ]", line)
                and extract_code_at_start(nxt)[0]
                and not extract_code_at_start(line)[0]
                and not any(r.match(line) for r in (PROF_RE, PREREQ_RE, DESC_RE))):
            continue

        # ── CLAVE: prefix (single-sheet format) ──────────────────────────────
        if CLAVE_RE.match(line):
            rest = CLAVE_RE.sub("", line).strip()
            code, _ = extract_code_at_start(rest)
            if code:
                flush()
                cur   = {"code": code, "title": "", "professor": "",
                         "prerequisites": "", "description": "",
                         "department": dept_from_code(code)}
                state = None
            continue

        # ── NOMBRE DEL CURSO ─────────────────────────────────────────────────
        if NOMBRE_RE.match(line) and cur:
            cur["title"] = NOMBRE_RE.sub("", line).strip().upper()
            state = None
            continue

        # ── Code at start ─────────────────────────────────────────────────────
        code, remainder = extract_code_at_start(line)
        if code and cur is not None:
            # A code line inside an open PROFESOR/PRERREQUISITOS block is a
            # prerequisite, unless a fresh PROFESOR label follows for a course
            # whose professor is already known (→ genuinely a new course).
            if state in ("prof", "prereq") and not (cur["professor"] and label_ahead(i)):
                p = cur["prerequisites"]
                if p == "Ninguno":
                    p = ""
                cur["prerequisites"] = (p + "\n" + line) if p else line
                state = "prereq"
                continue
            # Bare cross-listed code right after the title (e.g. EIN-xxxx / DER-xxxx)
            if (not remainder and state is None and cur["title"]
                    and not cur["professor"] and not cur["prerequisites"]):
                cur["title"] += f" ({code})"
                continue
        if code:
            flush()

            # remainder may be "TITLE PROFESOR: Name" or just "TITLE"
            # Split off any embedded PROFESOR field
            title_text = remainder
            embedded_prof = ""
            prof_split = re.split(r'PROFESOR(?:ES|AS|A)?\s*[:\.]?\s*', remainder, flags=re.IGNORECASE, maxsplit=1)
            if len(prof_split) == 2:
                title_text   = prof_split[0].strip()
                embedded_prof = prof_split[1].strip()

            cur = new_course(code, title_text.upper().strip().rstrip(".,"),
                             embedded_prof.rstrip(".,;"))
            state = "prof" if embedded_prof else None
            continue

        if cur is None:
            continue

        # ── PROFESOR ─────────────────────────────────────────────────────────
        if PROF_RE.match(line):
            val = PROF_RE.sub("", line).strip().rstrip(".,;")
            cur["professor"] = val
            state = "prof"
            continue

        # ── PRERREQUISITOS ────────────────────────────────────────────────────
        if PREREQ_RE.match(line):
            val = PREREQ_RE.sub("", line).strip().rstrip(".,;")
            if not val and re.match(r'^\s*SIN\s+PRERREQUISITO', line, re.IGNORECASE):
                val = "Ninguno"   # document explicitly states there are none
            cur["prerequisites"] = val
            state = "prereq"
            continue

        # ── DESCRIPCIÓN ──────────────────────────────────────────────────────
        if DESC_RE.match(line):
            val = DESC_RE.sub("", line).strip()
            desc_parts = [val] if val else []
            state = "desc"
            continue

        # ── Temario / notas / programa: stop collecting ──────────────────────
        if re.match(r'^\s*(TEMARIO|NOTAS?\s*:|PROGRAMA\s*:)', line, re.IGNORECASE):
            state = "ignore"
            continue
        if state == "ignore":
            continue

        # ── Continuation lines ────────────────────────────────────────────────
        is_prose = len(line) >= 90
        if state == "prof":
            if len(line) < 70:
                sep = " / " if cur["professor"] else ""
                cur["professor"] = (cur["professor"] + sep + line).strip(" /.,;")
            else:                                   # unlabeled description
                desc_parts.append(line); state = "desc"
        elif state == "prereq":
            p = cur["prerequisites"]
            if not p or not is_prose:
                if p == "Ninguno":
                    p = ""
                cur["prerequisites"] = (p + "\n" + line) if p else line
            else:                                   # unlabeled description
                desc_parts.append(line); state = "desc"
        elif state == "desc":
            desc_parts.append(line)
        elif state is None:
            if not cur["title"]:                     # title on line after code
                cur["title"] = line.upper().rstrip(".,")
            elif is_prose:                           # unlabeled description
                desc_parts.append(line); state = "desc"
            elif not cur["professor"] and not cur["prerequisites"]:
                sep = " " if line.startswith("(") else " — "         # subtitle
                cur["title"] += sep + line.upper().rstrip(".,")

    flush()
    return courses


# ─── EXCEL WRITER ─────────────────────────────────────────────────────────────

HEADER_FILL = PatternFill("solid", start_color="1F3864")
DEPT_FILL   = PatternFill("solid", start_color="2E75B6")
ALT_FILL    = PatternFill("solid", start_color="EBF3FA")
WHITE_FILL  = PatternFill("solid", start_color="FFFFFF")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
DEPT_FONT   = Font(name="Arial", bold=True, color="FFFFFF", size=10)
BODY_FONT   = Font(name="Arial", size=10)
_THIN       = Side(style="thin", color="B8CCE4")
CELL_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
WRAP_TOP    = Alignment(wrap_text=True, vertical="top")
CENTER_MID  = Alignment(horizontal="center", vertical="center", wrap_text=True)
COL_WIDTHS  = [22, 16, 44, 34, 48, 80]
HEADERS     = ["Departamento", "Código", "Título", "Profesor", "Prerrequisitos", "Descripción"]


def write_excel(courses: list[dict], output_path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Optativas Otoño 2026"

    for col_idx, width in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    for col_idx, header in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = HEADER_FONT;  cell.fill = HEADER_FILL
        cell.alignment = CENTER_MID;  cell.border = CELL_BORDER
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"

    def sort_key(c):
        dept = c["department"]
        return (DEPT_ORDER.index(dept) if dept in DEPT_ORDER else 99, c.get("code", ""))

    data_row = 2
    for course in sorted(courses, key=sort_key):
        fill = ALT_FILL if data_row % 2 == 1 else WHITE_FILL
        for col_idx, value in enumerate([
            course["department"], course.get("code",""), course.get("title",""),
            course.get("professor",""), course.get("prerequisites",""),
            course.get("description",""),
        ], 1):
            cell = ws.cell(row=data_row, column=col_idx, value=value or None)
            cell.font = BODY_FONT;  cell.fill = fill
            cell.alignment = WRAP_TOP;  cell.border = CELL_BORDER
        ws.row_dimensions[data_row].height = 90
        data_row += 1

    ws.auto_filter.ref = f"A1:F{data_row - 1}"
    wb.save(str(output_path))


# ─── FOLDER PICKER / .DOC CONVERSION / FILE FILTER ───────────────────────────

def pick_folder() -> Path:
    """Open a folder-selection dialog; fall back to ./input if no GUI is available."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        chosen = filedialog.askdirectory(title="Carpeta con archivos .doc / .docx")
        root.destroy()
        if chosen:
            return Path(chosen)
        sys.exit("No folder selected.")
    except SystemExit:
        raise
    except Exception:
        print("  (no GUI available — using ./input)")
        return Path("input")


def convert_doc(doc_path: Path, out_dir: Path) -> Path | None:
    """Convert a legacy .doc to .docx via Word (COM) or LibreOffice. None on failure."""
    target = out_dir / (doc_path.stem + ".docx")
    try:
        import win32com.client  # pywin32
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        try:
            d = word.Documents.Open(str(doc_path.resolve()), ReadOnly=True)
            d.SaveAs2(str(target.resolve()), FileFormat=16)
            d.Close(False)
        finally:
            word.Quit()
        if target.exists():
            return target
    except Exception:
        pass
    import shutil, subprocess
    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if soffice:
        try:
            subprocess.run([soffice, "--headless", "--convert-to", "docx",
                            "--outdir", str(out_dir), str(doc_path)],
                           check=True, capture_output=True, timeout=120)
            if target.exists():
                return target
        except Exception:
            pass
    return None


NOTES_RE = re.compile(r'^\s*NOTAS\s+AL\s+PLAN\s+DE\s+ESTUDIOS', re.IGNORECASE)

def is_plan_notes(doc) -> bool:
    """Study-plan notes files list codes/titles only — not lecture announcements."""
    for line in iter_lines(doc):
        return bool(NOTES_RE.match(line))
    return False


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def default_period(today=None):
    """Siguiente semestre a partir de hoy (igual que defaultPeriod() en la app web)."""
    import datetime
    d = today or datetime.date.today()
    if d.month <= 3:
        return "Verano", d.year
    if d.month <= 8:
        return "Otoño", d.year
    return "Primavera", d.year + 1


def main():
    parser = argparse.ArgumentParser(
        description="Parse .doc/.docx optativas in a folder → formatted Excel (no API needed)"
    )
    parser.add_argument("--input",   "-i", default=None,
                        help="Folder with .doc/.docx files (default: opens a folder picker)")
    def_sem, def_year = default_period()
    parser.add_argument("--semestre", "-s", default=def_sem, type=str.capitalize,
                        choices=["Primavera", "Verano", "Otoño", "Otono"],
                        help=f"Semestre para el nombre del archivo (default: {def_sem})")
    parser.add_argument("--anio", "-a", default=def_year, type=int,
                        help=f"Año para el nombre del archivo (default: {def_year})")
    parser.add_argument("--output",  "-o", default=None,
                        help="Output .xlsx path (default: Optativas_<Semestre>_<Año>.xlsx)")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Print every parsed course")
    args = parser.parse_args()
    if not args.output:
        sem = args.semestre.replace("ñ", "n")
        args.output = f"Optativas_{sem}_{args.anio}.xlsx"

    input_dir = (Path(args.input) if args.input else pick_folder()).resolve()
    if not input_dir.is_dir():
        sys.exit(f"Folder not found: {input_dir}")

    word_files = sorted(
        p for p in input_dir.iterdir()
        if p.is_file() and p.suffix.lower() in (".doc", ".docx")
        and not p.name.startswith("~$")
    )
    if not word_files:
        sys.exit(f"No .doc/.docx files found in: {input_dir}")

    print(f"📂  {input_dir}")
    print(f"    {len(word_files)} file(s) found\n")

    import tempfile
    tmp_dir = Path(tempfile.mkdtemp(prefix="optativas_"))

    all_courses: list[dict] = []
    seen_keys:   set[tuple] = set()
    duplicates:  list[str]  = []

    def norm(s: str) -> str:
        return re.sub(r"[^\w]", "", s or "").upper()

    for src_path in word_files:
        docx_path = src_path
        if src_path.suffix.lower() == ".doc":
            docx_path = convert_doc(src_path, tmp_dir)
            if docx_path is None:
                print(f"  ✗ {src_path.name}: could not convert .doc "
                      "(needs MS Word + pywin32, or LibreOffice)")
                continue
        try:
            if is_plan_notes(Document(str(docx_path))):
                print(f"  – {src_path.name:<50s}  skipped (notas al plan de estudios)")
                continue
            courses = parse_docx(docx_path)
        except Exception as e:
            print(f"  ✗ {src_path.name}: {e}")
            continue

        added = 0
        for c in courses:
            code = c.get("code", "").strip()
            if not code:
                continue
            key = (code, norm(c.get("title", "")))
            if key in seen_keys:
                duplicates.append(f"{code}  ← {src_path.name}")
            else:
                seen_keys.add(key)
                all_courses.append(c)
                added += 1
                if args.verbose:
                    print(f"      {code:16s}  {c['title'][:55]}")

        print(f"  {'✓' if added else '·'} {src_path.name:<50s}  {added} course(s)")

    print(f"\n{'─'*56}")
    print(f"Total courses : {len(all_courses)}")
    if duplicates:
        print(f"Duplicates    : {len(duplicates)} (skipped)")
        for d in duplicates:
            print(f"    {d}")

    if not all_courses:
        sys.exit("\nNo courses found — nothing to write.")

    dept_counts: dict[str, int] = {}
    for c in all_courses:
        dept_counts[c["department"]] = dept_counts.get(c["department"], 0) + 1

    print("\nBy department:")
    for dept in DEPT_ORDER:
        if dept in dept_counts:
            print(f"  {dept_counts[dept]:3d}  {dept}")
    for dept, cnt in dept_counts.items():
        if dept not in DEPT_ORDER:
            print(f"  {cnt:3d}  {dept}  ← unrecognised prefix")

    output_path = Path(args.output)
    write_excel(all_courses, output_path)
    print(f"\n✓  Excel written → {output_path.resolve()}")


if __name__ == "__main__":
    main()
