#!/usr/bin/env python3
"""Generates a Project Charter as XLSX. Usage: generate_charter.py [--lang en|pt-BR] [input.json] output.xlsx (no input: blank template)."""
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

LANGUAGES = ("en", "pt-BR")
LANG = "en"
PT_BR = {
    "TBD": "a definir",
    "Operations": "Operações",
    "Quality": "Qualidade",
    "Planning": "Planejamento",
    "Logistics": "Logística",
    "Safety": "Segurança",
    "Tax": "Fiscal",
    "Technology": "Tecnologia",
    "Purchasing": "Compras",
    "Sales": "Comercial",
    "Finance": "Financeiro",
    "People": "Pessoas",
    "Other": "Outros",
    "Approved": "Aprovado",
    "Rejected": "Reprovado",
    "Under review": "Em análise",
    "PROJECT CHARTER": "TERMO DE ABERTURA DO PROJETO",
    "PROJECT IDENTIFICATION": "IDENTIFICAÇÃO DO PROJETO",
    "Project title": "Título do projeto",
    "Code": "Código",
    "Start date": "Data de abertura",
    "Expected duration": "Duração prevista",
    "Project manager": "Gerente do projeto",
    "Owning area": "Área responsável",
    "1. BUSINESS CASE AND OBJECTIVE": "1. JUSTIFICATIVA E OBJETIVO",
    "Business case": "Justificativa (business case)",
    "Problem, opportunity or requirement that motivates the project.": "Problema, oportunidade ou exigência que motiva o projeto.",
    "Objective (SMART)": "Objetivo (SMART)",
    "Action verb + measurable result + deadline.": "Verbo no infinitivo + resultado mensurável + prazo.",
    "2. CURRENT PROCESS x FUTURE PROCESS": "2. PROCESSO ATUAL x PROCESSO FUTURO",
    "Current process (AS IS)": "Processo atual (AS IS)",
    "Future process (TO BE)": "Processo futuro (TO BE)",
    "How it works today, with its pains and limits.": "Como funciona hoje, com dores e limitações.",
    "How it will work after the project.": "Como passará a funcionar após o projeto.",
    "3. SCOPE": "3. ESCOPO",
    "Product, service or result scope": "Escopo do produto, serviço ou resultado",
    "Characteristics of the deliverable when it is done.": "Características do entregável quando pronto.",
    "In scope": "Contempla",
    "Out of scope": "Não contempla",
    "Main deliverables.": "Principais entregas.",
    "What will explicitly not be done.": "O que explicitamente não será feito.",
    "Acceptance criteria": "Critérios de aceitação",
    "Objective conditions to accept the deliverable.": "Condições objetivas para aceitar o entregável.",
    "Assumptions": "Premissas",
    "Constraints": "Restrições",
    "What is taken as true.": "Suposições tidas como verdadeiras.",
    "Limits of budget, time, technology and resources.": "Limites de orçamento, prazo, tecnologia e recursos.",
    "4. REQUIREMENTS": "4. REQUISITOS",
    "Functional requirement": "Requisito funcional",
    "Area": "Área",
    "Priority": "Prioridade",
    "Technical requirement": "Requisito técnico",
    "Owner": "Responsável",
    "Business rule": "Regra de negócio",
    "5. TECHNICAL SOLUTION": "5. SOLUÇÃO TÉCNICA",
    "Solution description": "Descrição da solução",
    "What will be implemented, developed or configured.": "O que será implementado, desenvolvido ou configurado.",
    "Systems involved": "Sistemas envolvidos",
    "Interfaces": "Interfaces",
    "Infrastructure": "Infraestrutura",
    "Source": "Origem",
    "Target": "Destino",
    "Type": "Tipo",
    "Integration description": "Descrição da integração",
    "6. OPERATIONAL IMPACTS": "6. IMPACTOS OPERACIONAIS",
    "7. RESOURCES": "7. RECURSOS",
    "Person or resource": "Responsável ou recurso",
    "Role and allocation": "Papel e alocação",
    "8. BUDGET": "8. ORÇAMENTO",
    "Cost category (excluding procurement)": "Categoria de custo (fora de aquisições)",
    "Amount": "Valor",
    "Total of the \"Procurement\" sheet": "Total da aba \"Aquisições\"",
    "TOTAL PLANNED BUDGET": "ORÇAMENTO TOTAL PREVISTO",
    "9. STATUS AND APPROVAL": "9. STATUS E APROVAÇÃO",
    "Role": "Papel",
    "Name and position": "Nome e cargo",
    "Date": "Data",
    "Signature": "Assinatura",
    "Notes": "Observações",
    "PROCUREMENT LIST": "RELAÇÃO DE AQUISIÇÕES",
    "Project": "Projeto",
    "TOTAL PLANNED AMOUNT": "VALOR TOTAL PREVISTO",
    "Item": "Objeto",
    "Qty": "Qtd",
    "Unit price": "Valor unitário",
    "Total": "Valor total",
    "Needed by": "Necessidade",
    "Charter": "TAP",
    "Procurement": "Aquisições",
    "Project Charter": "Termo de Abertura do Projeto"
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
                self.cell(1 + j * span, span, f"{mark}  {tr(option)}")
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
    s.cell(1, COLUMNS, tr("PROJECT CHARTER"), bold=True, color="FFFFFF", background=PRIMARY).font = Font(
        name=FONT, size=14, bold=True, color="FFFFFF")
    ws.row_dimensions[1].height = 24
    s.row += 1
    s.section(tr("PROJECT IDENTIFICATION"))
    for group in (
        ((tr("Project title"), "title"), (tr("Code"), "code"), (tr("Start date"), "start_date"), (tr("Expected duration"), "duration")),
        ((tr("Project manager"), "manager"), ("Sponsor", "sponsor"), (tr("Owning area"), "area"), ("PMO", "pmo")),
    ):
        s.labels([(label, 2) for label, _ in group])
        s.fields([(ident.get(key), 2, "") for _, key in group])

    s.section(tr("1. BUSINESS CASE AND OBJECTIVE"))
    s.labels([(tr("Business case"), COLUMNS)])
    s.text(d.get("business_case"), tr("Problem, opportunity or requirement that motivates the project."))
    s.labels([(tr("Objective (SMART)"), COLUMNS)])
    s.text(d.get("objective"), tr("Action verb + measurable result + deadline."))

    s.section(tr("2. CURRENT PROCESS x FUTURE PROCESS"))
    s.labels([(tr("Current process (AS IS)"), 4), (tr("Future process (TO BE)"), 4)])
    s.fields([(d.get("as_is"), 4, tr("How it works today, with its pains and limits.")),
              (d.get("to_be"), 4, tr("How it will work after the project."))])

    s.section(tr("3. SCOPE"))
    s.labels([(tr("Product, service or result scope"), COLUMNS)])
    s.text(d.get("product_scope"), tr("Characteristics of the deliverable when it is done."))
    s.labels([(tr("In scope"), 4), (tr("Out of scope"), 4)])
    s.fields([(d.get("in_scope"), 4, tr("Main deliverables.")), (d.get("out_of_scope"), 4, tr("What will explicitly not be done."))])
    s.labels([(tr("Acceptance criteria"), COLUMNS)])
    s.text(d.get("acceptance_criteria"), tr("Objective conditions to accept the deliverable."))
    s.labels([(tr("Assumptions"), 4), (tr("Constraints"), 4)])
    s.fields([(d.get("assumptions"), 4, tr("What is taken as true.")), (d.get("constraints"), 4, tr("Limits of budget, time, technology and resources."))])

    s.section(tr("4. REQUIREMENTS"))
    s.table([("ID", 1, "id"), (tr("Functional requirement"), 4, "description"), (tr("Area"), 2, "area"), (tr("Priority"), 1, "priority")], d.get("functional_requirements"))
    s.table([("ID", 1, "id"), (tr("Technical requirement"), 4, "description"), (tr("Owner"), 2, "owner"), (tr("Priority"), 1, "priority")], d.get("technical_requirements"))
    s.table([("ID", 1, "id"), (tr("Business rule"), 7, "description")], d.get("business_rules"))

    s.section(tr("5. TECHNICAL SOLUTION"))
    sol = d.get("solution", {})
    s.labels([(tr("Solution description"), COLUMNS)])
    s.text(sol.get("description"), tr("What will be implemented, developed or configured."))
    s.labels([(tr("Systems involved"), 3), (tr("Interfaces"), 3), (tr("Infrastructure"), 2)])
    s.fields([(sol.get("systems"), 3, ""), (sol.get("interfaces"), 3, ""), (sol.get("infrastructure"), 2, "")])
    s.table([(tr("Source"), 2, "source"), (tr("Target"), 2, "target"), (tr("Type"), 1, "type"), (tr("Integration description"), 3, "description")], d.get("integrations"))

    s.section(tr("6. OPERATIONAL IMPACTS"))
    s.checkboxes(AREAS, d.get("impacts") or [])

    s.section(tr("7. RESOURCES"))
    s.table([(tr("Person or resource"), 3, "name"), (tr("Area"), 2, "area"), (tr("Role and allocation"), 3, "role")], d.get("resources"))

    s.section(tr("8. BUDGET"))
    s.labels([(tr("Cost category (excluding procurement)"), 4), (tr("Type"), 2), (tr("Amount"), 2)])
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
    s.cell(1, 6, tr('Total of the "Procurement" sheet'), bold=True, background=LIGHT)
    s.cell(7, 2, f"='{tr('Procurement')}'!D4", bold=True, number_format=CURRENCY)
    s.row += 1
    s.cell(1, 6, tr("TOTAL PLANNED BUDGET"), bold=True, color="FFFFFF", background=PRIMARY)
    s.cell(7, 2, f"=G{s.row - 2}+G{s.row - 1}", bold=True, number_format=CURRENCY)
    s.row += 1

    s.section(tr("9. STATUS AND APPROVAL"))
    s.checkboxes(STATUS, [d.get("status")], per_row=3)
    roles = (tr("Project manager"), "Sponsor", "PMO")
    s.table([(tr("Role"), 2, "role"), (tr("Name and position"), 3, "name"), (tr("Date"), 1, "date"), (tr("Signature"), 2, "signature")],
            [{"role": r} for r in roles] if fill else [{}] * 3)
    if not fill:
        for i, r in enumerate(roles):
            ws.cell(s.row - 3 + i, 1, r)
    s.labels([(tr("Notes"), COLUMNS)])
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
    ws["A1"] = tr("PROCUREMENT LIST")
    ws["A1"].font, ws["A1"].fill = font(bold=True, color="FFFFFF"), dark
    ident = d.get("identification", {})
    ws["A2"], ws["B2"] = tr("Project"), ident.get("title") or (TBD if fill else None)
    ws["A3"], ws["B3"] = tr("Code"), ident.get("code") or (TBD if fill else None)
    ws.merge_cells("B2:H2")
    ws.merge_cells("B3:H3")
    ws["A2"].font = ws["A3"].font = font(bold=True)
    ws["B2"].font = ws["B3"].font = font()
    ws.merge_cells("A4:C4")
    ws["A4"] = tr("TOTAL PLANNED AMOUNT")
    ws["A4"].font, ws["A4"].fill = font(bold=True), light
    for i, h in enumerate([tr("Item"), tr("Qty"), tr("Unit price"), tr("Total"), tr("Type"), tr("Needed by"), tr("Owner"), "Status"], 1):
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
    wb.active.title = tr("Charter")
    procurement_sheet(wb.create_sheet(tr("Procurement")), d, fill, procurement)
    wb.properties.creator = wb.properties.lastModifiedBy = "pmo-ai-skills"
    wb.properties.title = tr("Project Charter")
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
    print(f"Charter: {args[1]} | planned budget: ${planned_total(data):,.2f}")


if __name__ == "__main__":
    main()
