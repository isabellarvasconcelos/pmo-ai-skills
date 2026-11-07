#!/usr/bin/env python3
"""Generates the lessons learned register as XLSX. Usage: generate_lessons.py [--lang en|pt-BR] [input.json] output.xlsx (no input: blank template)."""
import json
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

TBD = "TBD"
PRIMARY, LIGHT = "1F3A5F", "DCE6F1"
FONT = "Arial"
BORDER = Border(*[Side(style="thin", color="BFBFBF")] * 4)
TYPES = ["POSITIVE", "NEGATIVE"]
STATUS = ["OPEN", "IN PROGRESS", "DONE"]
TIMINGS = ["DURING THE PROJECT", "AT CLOSURE", "FUTURE PROJECTS"]
COLORS = {"POSITIVE": "C6E0B4", "NEGATIVE": "F4B0A5"}
FIRST_ROW = 10
SPARE_ROWS = 15
COLUMNS = [("ID", "id", 9), ("Date", "date", 11), ("Positive / negative", "type", 12), ("Description", "description", 44),
           ("Related area", "area", 15), ("Action plan", "action_plan", 40), ("Owner", "owner", 17),
           ("Closing date", "closing_date", 13), ("Status", "status", 14), ("Timing", "timing", 20),
           ("Notes", "notes", 28)]
OPTIONAL = {"closing_date", "notes"}  # left empty instead of "TBD"
ENUMS = {"type", "status", "timing"}  # written in the output language
ID_PREFIX = {"en": "LL", "pt-BR": "LIC"}

LANGUAGES = ("en", "pt-BR")
LANG = "en"
PT_BR = {
    "TBD": "a definir",
    "POSITIVE": "POSITIVO",
    "NEGATIVE": "NEGATIVO",
    "OPEN": "ABERTO",
    "IN PROGRESS": "EM ANDAMENTO",
    "DONE": "CONCLUÍDO",
    "DURING THE PROJECT": "DURANTE O PROJETO",
    "AT CLOSURE": "NO ENCERRAMENTO",
    "FUTURE PROJECTS": "PRÓXIMOS PROJETOS",
    "Date": "Data",
    "Positive / negative": "Positivo / negativo",
    "Description": "Detalhamento",
    "Related area": "Área relacionada",
    "Action plan": "Plano de ação",
    "Owner": "Responsável",
    "Closing date": "Data de encerramento",
    "Timing": "Momento da tratativa",
    "Notes": "Observações",
    "Lessons learned": "Lições aprendidas",
    "LESSONS LEARNED": "LIÇÕES APRENDIDAS",
    "Project code": "Código do projeto",
    "Project name": "Nome do projeto",
    "Positive": "Positivas",
    "Negative": "Negativas",
    "Open actions": "Ações abertas",
    "Lessons Learned": "Lições Aprendidas"
}


def tr(text):
    """Translates a fixed label to the output language. Input values stay as English codes."""
    return PT_BR.get(text, text) if LANG == "pt-BR" else text


def set_language(lang):
    global LANG, TBD
    if lang not in LANGUAGES:
        sys.exit(f"Invalid input:\n- language must be one of: {', '.join(LANGUAGES)}")
    LANG = lang
    TBD = tr("TBD")


def parse_args(argv):
    """Separates --lang from the positional arguments."""
    args, lang = list(argv), None
    if "--lang" in args:
        i = args.index("--lang")
        if i + 1 >= len(args):
            sys.exit(__doc__)
        lang = args[i + 1]
        del args[i:i + 2]
    return args, lang


def font(**k):
    return Font(name=FONT, size=10, **k)


def validate(d):
    errors, seen = [], set()
    for i, item in enumerate(d.get("lessons", []), 1):
        label = f"lessons[{i}]"
        if not item.get("description"):
            errors.append(f"{label}.description is required")
        for field, valid in (("type", TYPES), ("status", STATUS), ("timing", TIMINGS)):
            if item.get(field) is not None and item[field] not in valid:
                errors.append(f"{label}.{field} must be one of: {', '.join(valid)}")
        for field in ("date", "closing_date"):
            if item.get(field) is not None:
                try:
                    datetime.strptime(item[field], "%d/%m/%Y")
                except (TypeError, ValueError):
                    errors.append(f"{label}.{field} must use the DD/MM/YYYY format")
        if item.get("closing_date") and item.get("status") != "DONE":
            errors.append(f"{label}.closing_date only applies to status DONE")
        if item.get("id"):
            if item["id"] in seen:
                errors.append(f"{label}.id is duplicated: {item['id']}")
            seen.add(item["id"])
    if errors:
        sys.exit("Invalid input:\n- " + "\n- ".join(errors))


def number(items):
    """Keeps the given IDs and gives new items the next free number, without reusing gaps."""
    prefix = ID_PREFIX[LANG]
    size = len(prefix)
    numbers = [int(i["id"][size:]) for i in items if str(i.get("id", "")).startswith(prefix) and i["id"][size:].isdigit()]
    nxt = max(numbers, default=0) + 1
    result = []
    for item in items:
        item = dict(item)
        if not item.get("id"):
            item["id"] = f"{prefix}{nxt:03d}"
            nxt += 1
        result.append(item)
    return result, nxt


def generate(d, target, fill):
    wb = Workbook()
    ws = wb.active
    ws.title = tr("Lessons learned")
    last_col = chr(64 + len(COLUMNS))
    ws.merge_cells(f"A1:{last_col}1")
    ws["A1"] = tr("LESSONS LEARNED")
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=PRIMARY)
    ws.row_dimensions[1].height = 24

    ident = d.get("identification", {})
    for row, (label, key) in enumerate(((tr("Project code"), "code"), (tr("Project name"), "title"),
                                        (tr("Date"), "date"), (tr("Owner"), "owner")), 2):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
        ws.cell(row, 1, label).font = font(bold=True)
        ws.cell(row, 3, ident.get(key) or (TBD if fill else None)).font = font()

    items, nxt = number(d.get("lessons", []) if fill else [])
    last = FIRST_ROW + len(items) + SPARE_ROWS - 1
    summary = [(tr("Positive"), f'=COUNTIF(C{FIRST_ROW}:C{last},"{tr("POSITIVE")}")', sum(i.get("type") == "POSITIVE" for i in items)),
               (tr("Negative"), f'=COUNTIF(C{FIRST_ROW}:C{last},"{tr("NEGATIVE")}")', sum(i.get("type") == "NEGATIVE" for i in items)),
               (tr("Open actions"), f'=COUNTIF(I{FIRST_ROW}:I{last},"{tr("OPEN")}")+COUNTIF(I{FIRST_ROW}:I{last},"{tr("IN PROGRESS")}")',
                sum(i.get("status") in ("OPEN", "IN PROGRESS") for i in items))]
    computed = {}
    for row, (label, formula, value) in enumerate(summary, 2):
        ws.cell(row, 5, label).font = font(bold=True)
        c = ws.cell(row, 6, formula)
        c.alignment = Alignment(horizontal="left")
        c.font, c.fill, c.border = font(bold=True), PatternFill("solid", fgColor=LIGHT), BORDER
        computed[c.coordinate] = value

    header = FIRST_ROW - 1
    for j, (title, _, width) in enumerate(COLUMNS, 1):
        c = ws.cell(header, j, tr(title))
        c.font, c.border = font(bold=True, color="FFFFFF"), BORDER
        c.fill = PatternFill("solid", fgColor=PRIMARY)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[c.column_letter].width = width
    ws.row_dimensions[header].height = 32

    for i in range(len(items) + SPARE_ROWS):
        row = FIRST_ROW + i
        item = items[i] if i < len(items) else None
        for j, (_, key, width) in enumerate(COLUMNS, 1):
            if key == "id":
                value = item["id"] if item else f"{ID_PREFIX[LANG]}{nxt + i - len(items):03d}"
            elif item is None:
                value = None
            else:
                value = item.get(key) or (None if key in OPTIONAL else TBD)
                if key in ENUMS:
                    value = tr(value)
            c = ws.cell(row, j, value)
            c.font, c.border = font(), BORDER
            c.alignment = Alignment(wrap_text=True, vertical="top")
        if item:
            texts = [(item.get(k) or "", w) for _, k, w in COLUMNS if k in ("description", "action_plan", "notes")]
            ws.row_dimensions[row].height = 14 * max(len(t) // w + 1 for t, w in texts) + 2

    for column, options in (("C", TYPES), ("I", STATUS), ("J", TIMINGS)):
        dv = DataValidation(type="list", formula1='"' + ",".join(tr(o) for o in options) + '"', allow_blank=True)
        dv.add(f"{column}{FIRST_ROW}:{column}{last}")
        ws.add_data_validation(dv)
    for name, color in COLORS.items():
        ws.conditional_formatting.add(f"C{FIRST_ROW}:C{last}", CellIsRule(
            operator="equal", formula=[f'"{tr(name)}"'], fill=PatternFill("solid", bgColor=color)))
    ws.freeze_panes = ws.cell(FIRST_ROW, 4)
    ws.auto_filter.ref = f"A{header}:{last_col}{last}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    wb.properties.creator = wb.properties.lastModifiedBy = "pmo-ai-skills"
    wb.properties.title = tr("Lessons Learned")
    wb.calculation.fullCalcOnLoad = True
    wb.save(target)
    store_values(target, {1: computed})
    return items


def store_values(target, values):
    """Stores computed results next to the formulas.

    Excel recalculates on open, but previews (e-mail, Quick Look, browsers)
    only show the saved value. values: {sheet number: {"G10": value}}.
    """
    import re
    import zipfile
    from xml.sax.saxutils import escape

    with zipfile.ZipFile(target) as z:
        files = {n: z.read(n) for n in z.namelist()}
    for sheet, cells in values.items():
        name = f"xl/worksheets/sheet{sheet}.xml"
        xml = files[name].decode("utf-8")
        for ref, value in cells.items():
            if value is None or value == "":
                continue
            kind = ' t="str"' if isinstance(value, str) else ""
            content = escape(value) if isinstance(value, str) else repr(float(value))
            pattern = re.compile(rf'<c r="{ref}"( s="\d+")?><f>(.*?)</f>(?:<v\s*/>|<v></v>)</c>')
            xml, n = pattern.subn(lambda m: f'<c r="{ref}"{m.group(1) or ""}{kind}><f>{m.group(2)}</f><v>{content}</v></c>', xml)
            if n != 1:
                raise RuntimeError(f"Formula not found at {ref} on sheet {sheet}")
        files[name] = xml.encode("utf-8")
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for n, content in files.items():
            z.writestr(n, content)


def main():
    args, lang = parse_args(sys.argv[1:])
    if len(args) not in (1, 2):
        sys.exit(__doc__)
    if len(args) == 1:
        set_language(lang or "en")
        generate({}, args[0], fill=False)
        print(f"Blank template: {args[0]}")
        return
    data = json.loads(Path(args[0]).read_text(encoding="utf-8"))
    set_language(lang or data.get("language", "en"))
    validate(data)
    items = generate(data, args[1], fill=True)
    types = Counter(i.get("type") or TBD for i in items)
    open_items = [i["id"] for i in items if i.get("status") in ("OPEN", "IN PROGRESS")]
    print(f"Lessons: {args[1]} | {len(items)} lesson(s) | positive: {types['POSITIVE']}, negative: {types['NEGATIVE']}"
          + (f", no type: {types[TBD]}" if types[TBD] else "")
          + f" | open actions: {', '.join(open_items) or 'none'}")


if __name__ == "__main__":
    main()
