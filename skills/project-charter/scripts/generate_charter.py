#!/usr/bin/env python3
"""Generates a Project Charter as XLSX. Usage: generate_charter.py [input.json] output.xlsx (no input: blank template)."""
import json
import math
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col

TBD = "TBD"
PRIMARY, LIGHT, HINT = "1F3A5F", "DCE6F1", "7F7F7F"
FONT = "Arial"
COLUMN_WIDTH = 16
COLUMNS = 8
BORDER = Border(*[Side(style="thin", color="BFBFBF")] * 4)
CURRENCY = '"$"#,##0.00'
AREAS = ["Operations", "Quality", "Planning", "Logistics", "Safety", "Tax",
         "Technology", "Purchasing", "Sales", "Finance", "People", "Other"]
STATUS = ["Approved", "Rejected", "Under review"]


class Sheet:
    def __init__(self, ws, fill):
        self.ws, self.fill, self.row = ws, fill, 1
        for i in range(1, COLUMNS + 1):
            ws.column_dimensions[col(i)].width = COLUMN_WIDTH

    def cell(self, start, span, value, bold=False, color=None, background=None, hint=False, number_format=None):
        end = start + span - 1
        for c in range(start, end + 1):
            self.ws.cell(self.row, c).border = BORDER
        if span > 1:
            self.ws.merge_cells(start_row=self.row, start_column=start, end_row=self.row, end_column=end)
        cell = self.ws.cell(self.row, start, value)
        cell.font = Font(name=FONT, size=10, bold=bold, italic=hint, color=HINT if hint else color)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if background:
            cell.fill = PatternFill("solid", fgColor=background)
        if number_format:
            cell.number_format = number_format
        return cell

    def height(self, cells):
        """Merged cells do not resize on their own, so we estimate the height from the text."""
        lines = 1
        for text, span in cells:
            if isinstance(text, str) and not text.startswith("="):
                width = span * COLUMN_WIDTH * 1.05
                total = sum(max(1, math.ceil(len(p) / width)) for p in text.split("\n"))
                lines = max(lines, total)
        self.ws.row_dimensions[self.row].height = 14 * lines + 2

    def section(self, text):
        self.cell(1, COLUMNS, text, bold=True, color="FFFFFF", background=PRIMARY)
        self.row += 1

    def labels(self, items):
        start = 1
        for text, span in items:
            self.cell(start, span, text, bold=True, background=LIGHT)
            start += span
        self.height(items)
        self.row += 1

    def text(self, value, hint=""):
        self.fields([(value, COLUMNS, hint)])

    def fields(self, items):
        """items: (value, span, hint). In the blank template, the hint replaces the value."""
        start, sizes = 1, []
        for value, span, hint in items:
            if isinstance(value, list):
                value = "\n".join(f"• {v}" for v in value)
            if not self.fill:
                self.cell(start, span, hint or None, hint=True)
                sizes.append((hint, span))
            else:
                value = value if value not in (None, "", []) else TBD
                self.cell(start, span, value)
                sizes.append((value, span))
            start += span
        self.height(sizes)
        self.row += 1

    def table(self, columns, rows, minimum=3):
        """columns: (title, span, key)."""
        self.labels([(t, s) for t, s, _ in columns])
        if not self.fill:
            rows = [{}] * minimum
        elif not rows:
            rows = [{}]
        for item in rows:
            start, sizes = 1, []
            for _, span, key in columns:
                value = item.get(key) or (TBD if self.fill else None)
                self.cell(start, span, value)
                sizes.append((value, span))
                start += span
            self.height(sizes)
            self.row += 1

    def checkboxes(self, options, checked, per_row=4):
        span = COLUMNS // per_row
        for i in range(0, len(options), per_row):
            for j, option in enumerate(options[i:i + per_row]):
                mark = "☒" if option in checked else "☐"
                self.cell(1 + j * span, span, f"{mark}  {option}")
            self.row += 1


def validate(d):
    errors = []
    for group, field in (("costs", "amount"), ("procurement", "unit_price"), ("procurement", "quantity")):
        for i, item in enumerate(d.get(group, []), 1):
            v = item.get(field)
            if v is not None and (isinstance(v, bool) or not isinstance(v, (int, float)) or v < 0):
                errors.append(f"{group}[{i}].{field} must be a non-negative number")
    for group in ("costs", "procurement"):
        for i, item in enumerate(d.get(group, []), 1):
            if item.get("type") not in (None, "CAPEX", "OPEX"):
                errors.append(f"{group}[{i}].type must be CAPEX or OPEX")
    for key, valid in (("impacts", AREAS), ("status", STATUS)):
        values = d.get(key) or []
        for v in [values] if isinstance(values, str) else values:
            if v not in valid:
                errors.append(f"{key}: invalid value '{v}'. Use: {', '.join(valid)}")
    if errors:
        sys.exit("Invalid input:\n- " + "\n- ".join(errors))


def charter_sheet(ws, d, fill, computed):
    s = Sheet(ws, fill)
    ident = d.get("identification", {})
    s.cell(1, COLUMNS, "PROJECT CHARTER", bold=True, color="FFFFFF", background=PRIMARY).font = Font(
        name=FONT, size=14, bold=True, color="FFFFFF")
    ws.row_dimensions[1].height = 24
    s.row += 1
    s.section("PROJECT IDENTIFICATION")
    for group in (
        (("Project title", "title"), ("Code", "code"), ("Start date", "start_date"), ("Expected duration", "duration")),
        (("Project manager", "manager"), ("Sponsor", "sponsor"), ("Owning area", "area"), ("PMO", "pmo")),
    ):
        s.labels([(label, 2) for label, _ in group])
        s.fields([(ident.get(key), 2, "") for _, key in group])

    s.section("1. BUSINESS CASE AND OBJECTIVE")
    s.labels([("Business case", COLUMNS)])
    s.text(d.get("business_case"), "Problem, opportunity or requirement that motivates the project.")
    s.labels([("Objective (SMART)", COLUMNS)])
    s.text(d.get("objective"), "Action verb + measurable result + deadline.")

    s.section("2. CURRENT PROCESS x FUTURE PROCESS")
    s.labels([("Current process (AS IS)", 4), ("Future process (TO BE)", 4)])
    s.fields([(d.get("as_is"), 4, "How it works today, with its pains and limits."),
              (d.get("to_be"), 4, "How it will work after the project.")])

    s.section("3. SCOPE")
    s.labels([("Product, service or result scope", COLUMNS)])
    s.text(d.get("product_scope"), "Characteristics of the deliverable when it is done.")
    s.labels([("In scope", 4), ("Out of scope", 4)])
    s.fields([(d.get("in_scope"), 4, "Main deliverables."), (d.get("out_of_scope"), 4, "What will explicitly not be done.")])
    s.labels([("Acceptance criteria", COLUMNS)])
    s.text(d.get("acceptance_criteria"), "Objective conditions to accept the deliverable.")
    s.labels([("Assumptions", 4), ("Constraints", 4)])
    s.fields([(d.get("assumptions"), 4, "What is taken as true."), (d.get("constraints"), 4, "Limits of budget, time, technology and resources.")])

    s.section("4. REQUIREMENTS")
    s.table([("ID", 1, "id"), ("Functional requirement", 4, "description"), ("Area", 2, "area"), ("Priority", 1, "priority")], d.get("functional_requirements"))
    s.table([("ID", 1, "id"), ("Technical requirement", 4, "description"), ("Owner", 2, "owner"), ("Priority", 1, "priority")], d.get("technical_requirements"))
    s.table([("ID", 1, "id"), ("Business rule", 7, "description")], d.get("business_rules"))

    s.section("5. TECHNICAL SOLUTION")
    sol = d.get("solution", {})
    s.labels([("Solution description", COLUMNS)])
    s.text(sol.get("description"), "What will be implemented, developed or configured.")
    s.labels([("Systems involved", 3), ("Interfaces", 3), ("Infrastructure", 2)])
    s.fields([(sol.get("systems"), 3, ""), (sol.get("interfaces"), 3, ""), (sol.get("infrastructure"), 2, "")])
    s.table([("Source", 2, "source"), ("Target", 2, "target"), ("Type", 1, "type"), ("Integration description", 3, "description")], d.get("integrations"))

    s.section("6. OPERATIONAL IMPACTS")
    s.checkboxes(AREAS, d.get("impacts") or [])

    s.section("7. RESOURCES")
    s.table([("Person or resource", 3, "name"), ("Area", 2, "area"), ("Role and allocation", 3, "role")], d.get("resources"))

    s.section("8. BUDGET")
    s.labels([("Cost category (excluding procurement)", 4), ("Type", 2), ("Amount", 2)])
    costs = d.get("costs") or []
    first = s.row
    for item in costs if fill and costs else [{}] * (3 if not fill else 1):
        s.cell(1, 4, item.get("category") or (TBD if fill else None))
        s.cell(5, 2, item.get("type") or (TBD if fill else None))
        v = item.get("amount")
        s.cell(7, 2, v if v is not None else (TBD if fill else None), number_format=CURRENCY)
        s.row += 1
    last = s.row - 1
    s.cell(1, 6, "Subtotal", bold=True, background=LIGHT)
    s.cell(7, 2, f"=SUM(G{first}:G{last})", bold=True, number_format=CURRENCY)
    subtotal = sum(i.get("amount") or 0 for i in costs) if fill else 0
    purchases = sum((i.get("quantity") or 0) * (i.get("unit_price") or 0) for i in d.get("procurement", [])) if fill else 0
    computed.update({f"G{s.row}": subtotal, f"G{s.row + 1}": purchases, f"G{s.row + 2}": subtotal + purchases})
    s.row += 1
    s.cell(1, 6, 'Total of the "Procurement" sheet', bold=True, background=LIGHT)
    s.cell(7, 2, "='Procurement'!D4", bold=True, number_format=CURRENCY)
    s.row += 1
    s.cell(1, 6, "TOTAL PLANNED BUDGET", bold=True, color="FFFFFF", background=PRIMARY)
    s.cell(7, 2, f"=G{s.row - 2}+G{s.row - 1}", bold=True, number_format=CURRENCY)
    s.row += 1

    s.section("9. STATUS AND APPROVAL")
    s.checkboxes(STATUS, [d.get("status")], per_row=3)
    roles = ("Project manager", "Sponsor", "PMO")
    s.table([("Role", 2, "role"), ("Name and position", 3, "name"), ("Date", 1, "date"), ("Signature", 2, "signature")],
            [{"role": r} for r in roles] if fill else [{}] * 3)
    if not fill:
        for i, r in enumerate(roles):
            ws.cell(s.row - 3 + i, 1, r)
    s.labels([("Notes", COLUMNS)])
    s.text(d.get("notes"))

    ws.sheet_view.showGridLines = False
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def procurement_sheet(ws, d, fill, computed):
    widths = [38, 8, 16, 16, 9, 15, 20, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[col(i)].width = w
    dark = PatternFill("solid", fgColor=PRIMARY)
    light = PatternFill("solid", fgColor=LIGHT)
    font = lambda **k: Font(name=FONT, size=10, **k)
    ws.merge_cells("A1:H1")
    ws["A1"] = "PROCUREMENT LIST"
    ws["A1"].font, ws["A1"].fill = font(bold=True, color="FFFFFF"), dark
    ident = d.get("identification", {})
    ws["A2"], ws["B2"] = "Project", ident.get("title") or (TBD if fill else None)
    ws["A3"], ws["B3"] = "Code", ident.get("code") or (TBD if fill else None)
    ws.merge_cells("B2:H2")
    ws.merge_cells("B3:H3")
    ws["A2"].font = ws["A3"].font = font(bold=True)
    ws["B2"].font = ws["B3"].font = font()
    ws.merge_cells("A4:C4")
    ws["A4"] = "TOTAL PLANNED AMOUNT"
    ws["A4"].font, ws["A4"].fill = font(bold=True), light
    for i, h in enumerate(["Item", "Qty", "Unit price", "Total", "Type", "Needed by", "Owner", "Status"], 1):
        c = ws.cell(5, i, h)
        c.font, c.fill, c.border = font(bold=True), light, BORDER
    items = (d.get("procurement") or []) if fill else []
    n = max(len(items), 10)
    for i in range(n):
        row = 6 + i
        item = items[i] if i < len(items) else {}

        def other(key):
            return item.get(key) or (TBD if item else None)
        qty, unit = item.get("quantity"), item.get("unit_price")
        if qty is not None and unit is not None:
            computed[f"D{row}"] = qty * unit
        values = [other("item"), qty, unit, f'=IF(OR(B{row}="",C{row}=""),"",B{row}*C{row})',
                  other("type"), other("needed_by"), other("owner"), other("status")]
        for j, v in enumerate(values, 1):
            c = ws.cell(row, j, v)
            c.font, c.border = font(), BORDER
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if j in (3, 4):
                c.number_format = CURRENCY
    ws["D4"] = f"=SUM(D6:D{5 + n})"
    computed["D4"] = sum(v for k, v in computed.items() if k != "D4")
    ws["D4"].font, ws["D4"].number_format = font(bold=True), CURRENCY


def generate(d, target, fill):
    wb = Workbook()
    charter, procurement = {}, {}
    charter_sheet(wb.active, d, fill, charter)
    wb.active.title = "Charter"
    procurement_sheet(wb.create_sheet("Procurement"), d, fill, procurement)
    wb.properties.creator = wb.properties.lastModifiedBy = "pmo-ai-skills"
    wb.properties.title = "Project Charter"
    wb.calculation.fullCalcOnLoad = True
    wb.save(target)
    store_values(target, {1: charter, 2: procurement})


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


def planned_total(d):
    costs = sum(i.get("amount") or 0 for i in d.get("costs", []))
    purchases = sum((i.get("quantity") or 0) * (i.get("unit_price") or 0) for i in d.get("procurement", []))
    return costs + purchases


def main():
    if len(sys.argv) not in (2, 3):
        sys.exit(__doc__)
    if len(sys.argv) == 2:
        generate({}, sys.argv[1], fill=False)
        print(f"Blank template: {sys.argv[1]}")
        return
    data = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    validate(data)
    generate(data, sys.argv[2], fill=True)
    print(f"Charter: {sys.argv[2]} | planned budget: ${planned_total(data):,.2f}")


if __name__ == "__main__":
    main()
