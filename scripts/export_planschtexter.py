"""Exporterar befintliga och nya planschtexter från skolprogrammens Markdown-filer
till ett Word-dokument för kollegial granskning.

Körning (från projektroten):
    python scripts/export_planschtexter.py

Markdown-filerna läses bara, de ändras aldrig.
"""

import argparse
import difflib
import re
import sys
import unicodedata
from pathlib import Path

import yaml
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_INPUT = ROOT / "Texter"
DEFAULT_OUTPUT = ROOT / "exports" / "Navet_planschtexter_granskning.docx"

EXCLUDED_TITLES = []  # Lägg till titlar här för att utesluta program från exporten

EXTRA_ROWS = 5
FONT_NAME = "Nunito Light"
FONT_SIZE = Pt(11)
HEADER_FILL = "E7ECF3"
COLUMN_WIDTHS_CM = [6.675, 10.0125, 10.0125]  # 25 % / 37,5 % / 37,5 % av 26,7 cm
LONG_ROW_CHARS = 1500  # längre rader får delas mellan sidor
EXTRA_ROW_HEIGHT_CM = 3.5

EXISTING_HEADING = "Befintlig planschtext"
NEW_HEADING = "Ny planschtext"
WEB_HEADINGS = ["Ny webbtext", "Befintlig webbtext"]
HEADING_RE = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
DASHES_RE = re.compile(r"[\u2010-\u2015\u2212]")


class ExportError(Exception):
    pass


def normalize_title(title):
    title = unicodedata.normalize("NFC", title)
    title = DASHES_RE.sub("-", title)
    return re.sub(r"\s+", " ", title).strip().casefold()


def sort_key(title):
    # Svensk ordning: å, ä, ö efter z
    key = normalize_title(title)
    return key.replace("å", "{").replace("ä", "|").replace("ö", "}")


def read_file(path):
    try:
        return path.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        raise ExportError(f"Kunde inte läsa {path.name}: {exc}") from exc


def split_frontmatter(text, path):
    match = re.match(r"---\r?\n(.*?)\r?\n---[ \t]*(?:\r?\n|$)", text, re.S)
    if not match:
        return None, text
    try:
        data = yaml.safe_load(match.group(1))
    except yaml.YAMLError as exc:
        raise ExportError(f"Ogiltig YAML i {path.name}: {exc}") from exc
    return (data if isinstance(data, dict) else None), text[match.end():]


def extract_section(body, heading):
    """Returnerar rårader under rubriken (None om den saknas), fram till nästa rubrik på nivå 1-2."""
    wanted = normalize_title(heading)
    collected = None
    in_code = False
    for line in body.splitlines():
        if line.lstrip().startswith("```"):
            in_code = not in_code
        m = None if in_code else HEADING_RE.match(line)
        if collected is None:
            if m and len(m.group(1)) == 2 and normalize_title(m.group(2)) == wanted:
                collected = []
            continue
        if m and len(m.group(1)) <= 2:
            break
        collected.append(line)
    return collected


def clean_inline(text):
    text = re.sub(r"!\[([^\]]*)\]\([^)]*\)", r"\1", text)
    text = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"(\*\*|__)(.+?)\1", r"\2", text)
    text = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])", r"\1", text)
    text = re.sub(r"(?<![\w_])_(?!\s)(.+?)(?<!\s)_(?![\w_])", r"\1", text)
    text = re.sub(r"`([^`]*)`", r"\1", text)
    return text.strip()


def to_paragraphs(lines):
    """Gör om Markdown-rader till en lista med rena stycken."""
    if lines is None:
        return []
    paragraphs, current = [], []

    def flush():
        if current:
            paragraphs.append(" ".join(current))
            current.clear()

    for raw in lines:
        line = raw.rstrip()
        if not line.strip() or re.match(r"^\s*(-{3,}|\*{3,}|_{3,})\s*$", line):
            flush()
            continue
        line = re.sub(r"^\s*#{1,6}\s+", "", line)
        line = re.sub(r"^\s*>\s?", "", line)
        if re.match(r"^\s*[-*+]\s*$", line):
            flush()
            continue
        bullet = re.match(r"^\s*[-*+]\s+(.*)$", line)
        if bullet:
            flush()
            item = clean_inline(bullet.group(1))
            if item:
                paragraphs.append("\u2022 " + item)
            continue
        numbered = re.match(r"^\s*(\d+[.)])\s+(.*)$", line)
        if numbered:
            flush()
            paragraphs.append(f"{numbered.group(1)} {clean_inline(numbered.group(2))}")
            continue
        cleaned = clean_inline(line)
        if cleaned:
            current.append(cleaned)
    flush()
    return paragraphs


def load_programs(input_dir):
    files = sorted(input_dir.glob("*.md"))
    programs, skipped = [], []
    for path in files:
        text = read_file(path)
        meta, body = split_frontmatter(text, path)
        title = meta.get("titel") if meta else None
        if not title or not str(title).strip():
            skipped.append(path.name)
            continue
        programs.append(
            {
                "file": path.name,
                "title": re.sub(r"\s+", " ", str(title)).strip(),
                "existing": to_paragraphs(extract_section(body, EXISTING_HEADING)),
                "new": to_paragraphs(extract_section(body, NEW_HEADING)),
                "web": [to_paragraphs(extract_section(body, h)) for h in WEB_HEADINGS],
            }
        )
    return files, programs, skipped


def split_excluded(programs):
    wanted = {normalize_title(t): t for t in EXCLUDED_TITLES}
    kept, excluded, matched = [], [], set()
    for program in programs:
        key = normalize_title(program["title"])
        if key in wanted:
            excluded.append(program)
            matched.add(key)
        else:
            kept.append(program)
    unmatched = [t for k, t in wanted.items() if k not in matched]
    return kept, excluded, unmatched


def suggestions(title, programs):
    norm = {normalize_title(p["title"]): p["title"] for p in programs}
    key = normalize_title(title)
    contained = [t for n, t in norm.items() if n in key or key in n]
    close = [norm[c] for c in difflib.get_close_matches(key, list(norm), n=3, cutoff=0.5)]
    result = []
    for item in contained + close:
        if item not in result:
            result.append(item)
    return result[:3]


def set_font(font, element):
    font.name = FONT_NAME
    font.size = FONT_SIZE
    rpr = element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONT_NAME)


def add_text(paragraph, text, bold=False):
    run = paragraph.add_run(text)
    set_font(run.font, run._r)
    run.bold = bold
    return run


def fill_cell(cell, paragraphs, bold=False):
    first = True
    for text in paragraphs or [""]:
        para = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.space_after = Pt(4)
        add_text(para, text, bold=bold)


def shade(cell, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)


def row_property(row, tag, **attrs):
    el = OxmlElement(tag)
    for key, value in attrs.items():
        el.set(qn(f"w:{key}"), value)
    row._tr.get_or_add_trPr().append(el)


def configure_table(table):
    table.autofit = False
    tbl_pr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tbl_pr.append(layout)
    margins = OxmlElement("w:tblCellMar")
    for side, width in (("top", 80), ("left", 110), ("bottom", 80), ("right", 110)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(width))
        node.set(qn("w:type"), "dxa")
        margins.append(node)
    tbl_pr.append(margins)
    for col, width in zip(table.columns, COLUMN_WIDTHS_CM):
        col.width = Cm(width)


def set_widths(row):
    for cell, width in zip(row.cells, COLUMN_WIDTHS_CM):
        cell.width = Cm(width)


def add_page_number(section):
    para = section.footer.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_text(para, "Sida ")
    run = add_text(para, "")
    for kind, text in (("begin", None), (None, " PAGE "), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        run._r.append(el)


def build_document(programs, output):
    doc = Document()
    normal = doc.styles["Normal"]
    set_font(normal.font, normal.element)

    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = Cm(29.7), Cm(21.0)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, side, Cm(1.5))
    add_page_number(section)

    title = doc.add_paragraph()
    add_text(title, "Navets skolprogram \u2013 granskning av planschtexter", bold=True).font.size = Pt(18)
    intro = doc.add_paragraph()
    add_text(
        intro,
        "Jämför befintliga och föreslagna planschtexter. Lämna gärna synpunkter "
        "och ändringsförslag direkt i dokumentet.",
    )

    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    configure_table(table)

    header = table.rows[0]
    for cell, label in zip(header.cells, ["Skolprogram", "Befintlig planschtext", "Ny planschtext"]):
        fill_cell(cell, [label], bold=True)
        shade(cell, HEADER_FILL)
    set_widths(header)
    row_property(header, "w:cantSplit")
    row_property(header, "w:tblHeader")

    for program in programs:
        row = table.add_row()
        set_widths(row)
        fill_cell(row.cells[0], [program["title"]])
        fill_cell(row.cells[1], program["existing"])
        fill_cell(row.cells[2], program["new"])
        if sum(len(p) for p in program["existing"] + program["new"]) <= LONG_ROW_CHARS:
            row_property(row, "w:cantSplit")

    for number in range(1, EXTRA_ROWS + 1):
        row = table.add_row()
        set_widths(row)
        fill_cell(row.cells[0], [f"Nytt skolprogram {number}"])
        fill_cell(row.cells[1], [])
        fill_cell(row.cells[2], [])
        row_property(row, "w:cantSplit")
        row_property(row, "w:trHeight", val=str(int(EXTRA_ROW_HEIGHT_CM * 567)), hRule="atLeast")

    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        doc.save(output)
    except PermissionError as exc:
        raise ExportError(
            f"Kunde inte skriva {output}. Är filen öppen i Word? Stäng den och försök igen."
        ) from exc
    except OSError as exc:
        raise ExportError(f"Kunde inte skriva {output}: {exc}") from exc


def verify(output, exported, excluded):
    """Läser tillbaka dokumentet och kontrollerar innehållet. Returnerar felmeddelanden."""
    problems = []
    try:
        doc = Document(output)
    except Exception as exc:
        return [f"Dokumentet kunde inte öppnas igen: {exc}"]
    rows = doc.tables[0].rows[1:]
    if len(rows) != len(exported) + EXTRA_ROWS:
        problems.append(f"Fel antal rader: {len(rows)} (väntat {len(exported) + EXTRA_ROWS}).")
    titles = [r.cells[0].text for r in rows[: len(exported)]]
    if len(set(titles)) != len(titles):
        problems.append("Något skolprogram förekommer mer än en gång.")
    for program in excluded:
        if program["title"] in titles:
            problems.append(f"Exkluderat program finns med: {program['title']}")
    extra = rows[len(exported):]
    if [r.cells[0].text for r in extra] != [f"Nytt skolprogram {i}" for i in range(1, EXTRA_ROWS + 1)]:
        problems.append("De extra raderna är fel.")
    if any(r.cells[1].text or r.cells[2].text for r in extra):
        problems.append("De extra raderna ska ha tomma planschtextceller.")

    headings = {h.casefold() for h in WEB_HEADINGS + [EXISTING_HEADING, NEW_HEADING, "Innehåll", "Praktisk information"]}
    by_title = {p["title"]: p for p in exported}
    for row in rows[: len(exported)]:
        title = row.cells[0].text
        program = by_title.get(title)
        if program is None:
            problems.append(f"Okänd rad i dokumentet: {title}")
            continue
        for cell, expected in zip(row.cells[1:], (program["existing"], program["new"])):
            if [p.text for p in cell.paragraphs if p.text] != expected:
                problems.append(f"Cellinnehållet avviker från källan för {title}.")
            for para in cell.paragraphs:
                if para.text.strip().casefold() in headings or para.text.lstrip().startswith("#"):
                    problems.append(f"Rubrik följde med för {title}: {para.text!r}")
    return problems


def identical_to_web(programs):
    """Planschtexter som i källfilen råkar vara ordagrant lika en webbtext (inte ett fel)."""
    return [
        p["title"]
        for p in programs
        if any(text and text in p["web"] for text in (p["existing"], p["new"]))
    ]


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="Mapp med Markdown-filer")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Sökväg till .docx")
    args = parser.parse_args()

    if not args.input.is_dir():
        raise ExportError(f"Hittar inte mappen med Markdown-filer: {args.input}")

    files, programs, skipped = load_programs(args.input)
    kept, excluded, unmatched = split_excluded(programs)
    kept.sort(key=lambda p: sort_key(p["title"]))
    build_document(kept, args.output)

    print(f"Markdown-filer hittade:        {len(files)}")
    if skipped:
        print(f"  Utan YAML-titel (hoppades över): {', '.join(skipped)}")
    print(f"Skolprogram exporterade:       {len(kept)}")
    if EXCLUDED_TITLES:
        print(f"Exkluderade program ({len(excluded)} av {len(EXCLUDED_TITLES)}):")
        for p in excluded:
            print(f"  - {p['title']}  [{p['file']}]")
    else:
        print("Exkluderade program: inga (alla skolprogram exporteras)")

    if unmatched:
        print("Titlar i exkluderingslistan som INTE kunde matchas:")
        for title in unmatched:
            print(f"  - {title}")
            hint = suggestions(title, kept)
            if hint:
                print(f"      liknande titlar som fortfarande är med: {' | '.join(hint)}")

    missing = [
        (p["title"], [n for n, v in (("befintlig", p["existing"]), ("ny", p["new"])) if not v])
        for p in kept
        if not p["existing"] or not p["new"]
    ]
    if missing:
        print("Program med saknad/tom planschtext:")
        for title, what in missing:
            print(f"  - {title}: saknar {' och '.join(what)}")

    print(f"Extra rader för nya skolprogram: {EXTRA_ROWS} tillagda")
    problems = verify(args.output, kept, excluded)
    print("Kontroll av dokumentet: " + ("OK" if not problems else "FEL"))
    for problem in problems:
        print(f"  ! {problem}")
    same = identical_to_web(kept)
    if same:
        print("Obs: planschtexten är ordagrant lika webbtexten i källfilen (exporterad från planschtext-rubriken):")
        for title in same:
            print(f"  - {title}")
    print(f"Dokument: {args.output}")
    return 1 if problems else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except ExportError as exc:
        print(f"FEL: {exc}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        print(f"Exporten misslyckades: {type(exc).__name__}: {exc}", file=sys.stderr)
        sys.exit(1)
