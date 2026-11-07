#!/usr/bin/env python3
"""Generates the risk matrix as XLSX. Usage: generate_matrix.py [--lang en|pt-BR] [input.json] output.xlsx (no input: blank template)."""
import json
import sys
from collections import Counter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

TBD = "TBD"
PRIMARY, LIGHT = "1F3A5F", "DCE6F1"
FONT = "Arial"
BORDER = Border(*[Side(style="thin", color="BFBFBF")] * 4)
CURRENCY = '"$"#,##0.00'
SIGNIFICANT, IMPORTANT = 15, 5  # minimum score of each rating
COLORS = {"Not significant": "C6E0B4", "Important": "FFE699", "Significant": "F4B0A5"}
TYPES = ["RISK", "ISSUE", "OPPORTUNITY"]
STATUS = ["OPEN", "CLOSED", "CANCELLED"]
THREAT_RESPONSES = ["Mitigate", "Share", "Transfer", "Avoid", "Accept"]
OPPORTUNITY_RESPONSES = ["Exploit", "Share", "Enhance", "Accept"]
RESPONSES = THREAT_RESPONSES + [r for r in OPPORTUNITY_RESPONSES if r not in THREAT_RESPONSES]
SCALE = ["1", "2", "3", "4", "5"]
FIRST_ROW = 8
ID_PREFIX = {"en": "RISK", "pt-BR": "RISC"}
SPARE_ROWS = 15
COLUMNS = [("ID", 10), ("Description", 42), ("Type", 15), ("Impact", 34), ("Impact level (1-5)", 10),
           ("Probability (1-5)", 11), ("Rating", 17), ("Response", 13), ("Action plan", 42),
           ("Cost", 14), ("Status", 13), ("Notes", 30)]

LANGUAGES = ("en", "pt-BR")
LANG = "en"
PT_BR = {
    "TBD": "a definir",
    "Not significant": "Não Significativo",
    "Important": "Importante",
    "Significant": "Significativo",
    "RISK": "RISCO",
    "ISSUE": "ISSUE",
    "OPPORTUNITY": "OPORTUNIDADE",
    "OPEN": "ABERTO",
    "CLOSED": "CONCLUÍDO",
    "CANCELLED": "CANCELADO",
    "Mitigate": "Mitigar",
    "Share": "Compartilhar",
    "Transfer": "Transferir",
    "Avoid": "Eliminar",
    "Accept": "Aceitar",
    "Exploit": "Explorar",
    "Enhance": "Melhorar",
    "Description": "Descrição",
    "Type": "Tipo",
    "Impact": "Consequência",
    "Impact level (1-5)": "Nível (1-5)",
    "Probability (1-5)": "Probabilidade (1-5)",
    "Rating": "Grau",
    "Response": "Medida",
    "Action plan": "Plano de ação",
    "Cost": "Custo",
    "Notes": "Observações",
    "RISK, ISSUE AND OPPORTUNITY MATRIX": "MATRIZ DE RISCOS, ISSUES E OPORTUNIDADES",
    "Project code": "Código do projeto",
    "Project name": "Nome do projeto",
    "Date": "Data",
    "Owner": "Responsável",
    "RATING = IMPACT x PROBABILITY": "GRAU = CONSEQUÊNCIA x PROBABILIDADE",
    "Impact / Probability": "Consequência / Probabilidade",
    "Matrix": "Matriz",
    "Risk, Issue and Opportunity Matrix": "Matriz de Riscos, Issues e Oportunidades"
}


def tr(text):
    """Translates a fixed label to the output language. Input values stay as English codes."""
    return PT_BR.get(text, text) if LANG == "pt-BR" else text


def set_language(lang):
    global LANG, TBD, CURRENCY
    if lang not in LANGUAGES:
        sys.exit(f"Invalid input:\n- language must be one of: {', '.join(LANGUAGES)}")
    LANG = lang
    TBD = tr("TBD")
    CURRENCY = '"R$" #,##0.00' if lang == "pt-BR" else '"$"#,##0.00'


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


def rating(score):
    if score >= SIGNIFICANT:
        return "Significant"
    return "Important" if score >= IMPORTANT else "Not significant"


def rating_formula(ref):
    return f'IF({ref}>={SIGNIFICANT},"{tr("Significant")}",IF({ref}>={IMPORTANT},"{tr("Important")}","{tr("Not significant")}"))'


def font(**k):
    return Font(name=FONT, size=10, **k)


def validate(d):
    errors, seen = [], set()
    for i, item in enumerate(d.get("items", []), 1):
        label = f"items[{i}]"
        if not item.get("description"):
            errors.append(f"{label}.description is required")
        if item.get("type") not in TYPES:
            errors.append(f"{label}.type must be one of: {', '.join(TYPES)}")
        for field in ("impact_level", "probability"):
            v = item.get(field)
            if v is not None and (type(v) is not int or not 1 <= v <= 5):
                errors.append(f"{label}.{field} must be an integer from 1 to 5")
        allowed = OPPORTUNITY_RESPONSES if item.get("type") == "OPPORTUNITY" else THREAT_RESPONSES
        if item.get("response") is not None and item["response"] not in allowed:
            errors.append(f"{label}.response must be one of: {', '.join(allowed)}")
        if item.get("status") is not None and item["status"] not in STATUS:
            errors.append(f"{label}.status must be one of: {', '.join(STATUS)}")
        c = item.get("cost")
        if c is not None and (type(c) not in (int, float) or c < 0):
            errors.append(f"{label}.cost must be a non-negative number")
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


def matrix_sheet(ws, d, fill, computed):
    ident = d.get("identification", {})
    ws.merge_cells("A1:L1")
    ws["A1"] = tr("RISK, ISSUE AND OPPORTUNITY MATRIX")
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=PRIMARY)
    ws.row_dimensions[1].height = 24
    for row, (label, key) in enumerate(((tr("Project code"), "code"), (tr("Project name"), "title"), (tr("Date"), "date"), (tr("Owner"), "owner")), 2):
        ws.cell(row, 1, label).font = font(bold=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
        ws.cell(row, 3, ident.get(key) or (TBD if fill else None)).font = font()

    header = FIRST_ROW - 1
    for j, (title, width) in enumerate(COLUMNS, 1):
        c = ws.cell(header, j, tr(title))
        c.font, c.border = font(bold=True, color="FFFFFF"), BORDER
        c.fill = PatternFill("solid", fgColor=PRIMARY)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[c.column_letter].width = width
    ws.row_dimensions[header].height = 32

    items, nxt = number(d.get("items", []) if fill else [])
    for i in range(len(items) + SPARE_ROWS):
        row = FIRST_ROW + i
        item = items[i] if i < len(items) else None
        item_id = item["id"] if item else f"{ID_PREFIX[LANG]}{nxt + i - len(items):03d}"
        empty = TBD if item else None
        get = (lambda k: item.get(k) if item.get(k) is not None else empty) if item else (lambda k: None)
        score = f"E{row}*F{row}"
        if item:
            c, p = item.get("impact_level"), item.get("probability")
            computed[f"G{row}"] = tr(rating(c * p)) if c and p else TBD
        values = [item_id, get("description"), tr(get("type")), get("impact"), get("impact_level"), get("probability"),
                  f'=IF(COUNT(E{row},F{row})=2,{rating_formula(score)},IF(B{row}="","","{TBD}"))',
                  tr(get("response")), get("action_plan"), get("cost"), tr(get("status")), item.get("notes") if item else None]
        for j, v in enumerate(values, 1):
            c = ws.cell(row, j, v)
            c.font, c.border = font(), BORDER
            c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row, 10).number_format = CURRENCY
        if item:
            texts = [(item.get("description"), 42), (item.get("impact"), 34), (item.get("action_plan"), 42), (item.get("notes"), 30)]
            ws.row_dimensions[row].height = 14 * max(1, *[len(t or "") // w + 1 for t, w in texts]) + 2
    last = FIRST_ROW + len(items) + SPARE_ROWS - 1

    for column, options in (("C", TYPES), ("E", SCALE), ("F", SCALE), ("H", RESPONSES), ("K", STATUS)):
        dv = DataValidation(type="list", formula1='"' + ",".join(tr(o) for o in options) + '"', allow_blank=True)
        dv.add(f"{column}{FIRST_ROW}:{column}{last}")
        ws.add_data_validation(dv)
    for name, color in COLORS.items():
        ws.conditional_formatting.add(f"G{FIRST_ROW}:G{last}", CellIsRule(
            operator="equal", formula=[f'"{tr(name)}"'], fill=PatternFill("solid", bgColor=color)))
    ws.freeze_panes = ws.cell(FIRST_ROW, 3)
    ws.auto_filter.ref = f"A{header}:L{last}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def rating_sheet(ws):
    ws.merge_cells("A1:F1")
    ws["A1"] = tr("RATING = IMPACT x PROBABILITY")
    ws["A1"].font = font(bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=PRIMARY)
    ws["A2"] = tr("Impact / Probability")
    ws["A2"].font, ws["A2"].alignment = font(bold=True), Alignment(wrap_text=True)
    ws.column_dimensions["A"].width = 18
    for n in range(1, 6):
        ws.column_dimensions[chr(65 + n)].width = 20
        for c in (ws.cell(2, n + 1, n), ws.cell(n + 2, 1, n)):
            c.font, c.fill, c.border = font(bold=True), PatternFill("solid", fgColor=LIGHT), BORDER
    for imp in range(1, 6):
        for prob in range(1, 6):
            label = rating(imp * prob)
            c = ws.cell(imp + 2, prob + 1, f"{imp * prob} - {tr(label)}")
            c.font, c.border = font(), BORDER
            c.fill = PatternFill("solid", fgColor=COLORS[label])


def generate(d, target, fill):
    wb = Workbook()
    wb.active.title = tr("Matrix")
    computed = {}
    matrix_sheet(wb.active, d, fill, computed)
    rating_sheet(wb.create_sheet(tr("Rating")))
    wb.properties.creator = wb.properties.lastModifiedBy = "pmo-ai-skills"
    wb.properties.title = tr("Risk, Issue and Opportunity Matrix")
    wb.calculation.fullCalcOnLoad = True
    wb.save(target)
    store_values(target, {1: computed})


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
    generate(data, args[1], fill=True)
    items = data.get("items", [])
    ratings = Counter(
        rating(i["impact_level"] * i["probability"]) if i.get("impact_level") and i.get("probability") else TBD
        for i in items)
    summary = ", ".join(f"{n}: {ratings[n]}" for n in [*COLORS, TBD] if ratings[n])
    print(f"Matrix: {args[1]} | {len(items)} items | {summary or 'no items'}")


if __name__ == "__main__":
    main()
