#!/usr/bin/env python3
"""Generates the project closure as DOCX. Usage: generate_closure.py [--lang en|pt-BR] [input.json] output.docx (no input: blank template)."""
import json
import sys
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

TBD = "TBD"
FONT = "Arial"
PRIMARY, LIGHT, HINT = "1F3A5F", "DCE6F1", "7F7F7F"
REASONS = ["COMPLETION", "REPLACEMENT", "CANCELLATION"]
WIDTH = 17.0  # usable cm on A4 with 2 cm margins
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW",
               "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"]

LANGUAGES = ("en", "pt-BR")
LANG = "en"
PT_BR = {
    "TBD": "a definir",
    "Completion": "Conclusão",
    "Replacement": "Substituição",
    "Cancellation": "Cancelamento",
    "PROJECT CLOSURE": "TERMO DE ENCERRAMENTO DO PROJETO",
    "Project code": "Código do projeto",
    "Project name": "Nome do projeto",
    "Date": "Data",
    "Owner": "Responsável",
    "Purpose of this document": "Objetivo do documento",
    "[project name]": "[nome do projeto]",
    "This document formalizes the closure of the {} project, consolidating the results achieved, the main lessons learned, the deliverables and the comparison between plan and execution. With the deliverables validated, the project team's work and its responsibilities for the defined scope come to an end.": "Este documento formaliza o encerramento do projeto {}, consolidando os resultados obtidos, os principais aprendizados, as entregas realizadas e a comparação entre o planejado e o executado. Com a validação das entregas, encerra-se a atuação da equipe do projeto e as responsabilidades relacionadas ao escopo definido.",
    "Lessons learned": "Aprendizados do projeto",
    "Lesson learned": "Lição aprendida",
    "Action plan": "Plano de ação",
    "Planned x actual": "Comparativo previsto x realizado",
    "Indicator": "Indicador",
    "Original plan": "Previsto original",
    "Approved changes": "Mudanças aprovadas",
    "Updated plan": "Previsto atualizado",
    "Actual": "Realizado",
    "Deviation": "Desvio",
    "{:+d} days": "{:+d} dias",
    "{} days": "{} dias",
    "Schedule": "Prazo",
    "Cost": "Custo",
    "None provided": "Nenhuma informada",
    "Approved changes considered: ": "Mudanças aprovadas consideradas: ",
    "Scope status": "Status do escopo",
    "Planned scope completed?": "Escopo planejado concluído?",
    "If not, explain": "Caso não, justificar",
    "Deliverables accepted?": "Entregas homologadas?",
    "Open items remaining?": "Pendências remanescentes?",
    "If yes, describe": "Caso sim, descrever",
    "☐ Yes  ☐ No": "☐ Sim  ☐ Não",
    "Yes": "Sim",
    "No": "Não",
    "Closure reason": "Motivo do encerramento",
    "name the project that replaces it": "indicar o projeto que o substitui",
    "describe the reason": "descrever o motivo",
    "The deliverables in the project scope are hereby declared completed, and the applicable documentation has been prepared, stored and made available to the responsible areas.": "Declara-se que as entregas previstas no escopo do projeto foram concluídas e que a documentação aplicável foi elaborada, armazenada e disponibilizada às áreas responsáveis.",
    "Final remarks": "Considerações finais",
    "Signatures": "Assinaturas",
    "Project manager": "Gerente do projeto",
    "Role": "Papel",
    "Name": "Nome",
    "Signature": "Assinatura",
    "Project Closure": "Termo de Encerramento do Projeto"
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


def element(tag, **attributes):
    el = OxmlElement(f"w:{tag}")
    for k, v in attributes.items():
        el.set(qn(f"w:{k}"), v)
    return el


def shade(cell, fill):
    cell._tc.get_or_add_tcPr().append(element("shd", val="clear", color="auto", fill=fill))


def style_run(run, bold=False, color=None, italic=False, size=10):
    run.font.name, run.font.size, run.bold, run.italic = FONT, Pt(size), bold, italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def write(cell, text, bold=False, color=None, hint=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    style_run(p.add_run(str(text)), bold, HINT if hint else color, italic=hint)


def table(doc, rows, widths):
    """Table with fixed width and explicit borders, so it does not depend on the reader's styles."""
    t = doc.add_table(rows=rows, cols=len(widths))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblPr = t._tbl.tblPr
    borders = element("tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(element(side, val="single", sz="4", space="0", color="BFBFBF"))
    children = {el.tag.split("}")[1]: el for el in tblPr}
    children.update({"tblW": element("tblW", w=str(int(sum(widths) * 567)), type="dxa"),
                     "tblLayout": element("tblLayout", type="fixed"), "tblBorders": borders})
    for el in list(tblPr):
        tblPr.remove(el)
    for tag in TBLPR_ORDER:  # Word requires this order
        if tag in children:
            tblPr.append(children[tag])
    for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(int(w * 567)))
    for j, w in enumerate(widths):
        for cell in t.columns[j].cells:
            cell.width = Cm(w)
    return t


def header_row(t, titles):
    for j, title in enumerate(titles):
        write(t.cell(0, j), title, bold=True)
        shade(t.cell(0, j), LIGHT)


def section(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(10), Pt(4)
    style_run(p.add_run(title.upper()), bold=True, color=PRIMARY, size=11)
    border = element("pBdr")
    border.append(element("bottom", val="single", sz="8", space="1", color=PRIMARY))
    p._p.get_or_add_pPr().append(border)


def paragraph(doc, text, hint=False, bold=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    style_run(p.add_run(text), bold=bold, color=HINT if hint else None, italic=hint)


def date(text):
    return datetime.strptime(text, "%d/%m/%Y")


def money(v, sign=False):
    """Formats an amount as $1,234.56 or R$ 1.234,56."""
    amount = f"{abs(v):,.2f}"
    text = "R$ " + amount.translate(str.maketrans(",.", ".,")) if LANG == "pt-BR" else "$" + amount
    if sign:
        return ("+" if v >= 0 else "-") + text
    return text if v >= 0 else "-" + text


def pct(v):
    text = f"{v:+.1f}%"
    return text.replace(".", ",") if LANG == "pt-BR" else text


def days(v, sign=False):
    return tr("{:+d} days").format(v) if sign else tr("{} days").format(v)


def comparison(d):
    """Computes schedule (calendar days) and cost: original, changes, updated, actual and deviation."""
    base, actual = d.get("baseline", {}), d.get("actual", {})
    changes = d.get("approved_changes", [])
    days_changed = sum(c.get("schedule_days") or 0 for c in changes)
    cost_changed = sum(c.get("cost") or 0 for c in changes)
    r = {"schedule": {}, "cost": {}}
    if base.get("start_date") and base.get("end_date"):
        original = (date(base["end_date"]) - date(base["start_date"])).days
        r["schedule"] = {"original": original, "changes": days_changed, "updated": original + days_changed}
        if actual.get("end_date"):
            r["schedule"]["actual"] = (date(actual["end_date"]) - date(base["start_date"])).days
    if base.get("cost") is not None:
        r["cost"] = {"original": base["cost"], "changes": cost_changed, "updated": base["cost"] + cost_changed}
        if actual.get("cost") is not None:
            r["cost"]["actual"] = actual["cost"]
    for x in r.values():
        if "actual" in x:
            x["deviation"] = x["actual"] - x["updated"]
            x["deviation_pct"] = x["deviation"] / x["updated"] * 100 if x["updated"] else None
    return r


def generate(d, target, fill):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(2)
    for style in ("Normal", "List Bullet"):
        doc.styles[style].font.name = FONT
        doc.styles[style].font.size = Pt(10)

    def v(value):
        if not fill:
            return ""
        return value if value not in (None, "", []) else TBD

    style_run(doc.add_paragraph().add_run(tr("PROJECT CLOSURE")), bold=True, color=PRIMARY, size=16)

    ident = d.get("identification", {})
    fields = [(tr("Project code"), "code"), (tr("Project name"), "title"), (tr("Date"), "date"), (tr("Owner"), "owner")]
    t = table(doc, len(fields), [4.5, WIDTH - 4.5])
    for i, (label, key) in enumerate(fields):
        write(t.cell(i, 0), label, bold=True)
        shade(t.cell(i, 0), LIGHT)
        write(t.cell(i, 1), v(ident.get(key)))

    section(doc, tr("Purpose of this document"))
    name = ident.get("title") if fill and ident.get("title") else tr("[project name]")
    paragraph(doc, tr("This document formalizes the closure of the {} project, consolidating the results achieved, "
                      "the main lessons learned, the deliverables and the comparison between plan and execution. "
                      "With the deliverables validated, the project team's work and its responsibilities for the "
                      "defined scope come to an end.").format(name))

    section(doc, tr("Lessons learned"))
    lessons = d.get("lessons", []) if fill else []
    rows = lessons or ([{}] if fill else [{}] * 3)
    t = table(doc, len(rows) + 1, [1.7, 7.9, 7.4])
    header_row(t, ["ID", tr("Lesson learned"), tr("Action plan")])
    for i, item in enumerate(rows, 1):
        write(t.cell(i, 0), item.get("id") or (f"{'LIC' if LANG == 'pt-BR' else 'LL'}{i:03d}" if item else ""))
        write(t.cell(i, 1), v(item.get("lesson")))
        write(t.cell(i, 2), v(item.get("action_plan")))

    section(doc, tr("Planned x actual"))
    comp = comparison(d) if fill else {"schedule": {}, "cost": {}}
    t = table(doc, 3, [2.3, 2.9, 2.6, 2.9, 3.0, 3.3])
    header_row(t, [tr("Indicator"), tr("Original plan"), tr("Approved changes"), tr("Updated plan"), tr("Actual"), tr("Deviation")])
    formats = {"schedule": days, "cost": money}
    for i, (label, key) in enumerate(((tr("Schedule"), "schedule"), (tr("Cost"), "cost")), 1):
        write(t.cell(i, 0), label, bold=True)
        x, fmt = comp[key], formats[key]
        for j, field in enumerate(("original", "changes", "updated", "actual"), 1):
            if field == "changes" and field in x and not d.get("approved_changes"):
                write(t.cell(i, j), tr("None provided"), hint=True)
            else:
                write(t.cell(i, j), fmt(x[field], field == "changes") if field in x else v(None))
        if "deviation" in x:
            text = fmt(x["deviation"], True) + (f" ({pct(x['deviation_pct'])})" if x["deviation_pct"] is not None else "")
            write(t.cell(i, 5), text, bold=True)
        else:
            write(t.cell(i, 5), v(None))
    changes = d.get("approved_changes", []) if fill else []
    if changes:
        paragraph(doc, tr("Approved changes considered: ") + "; ".join(
            f"{c.get('code', TBD)} ({days(c.get('schedule_days') or 0, True)}, {money(c.get('cost') or 0, True)})" for c in changes) + ".", hint=True)

    section(doc, tr("Scope status"))
    scope = d.get("scope", {})
    questions = [(tr("Planned scope completed?"), "completed", tr("If not, explain")),
                 (tr("Deliverables accepted?"), "accepted", tr("If not, explain")),
                 (tr("Open items remaining?"), "open_items", tr("If yes, describe"))]
    t = table(doc, 3, [5.0, 2.6, 9.4])
    for i, (label, key, hint) in enumerate(questions):
        write(t.cell(i, 0), label, bold=True)
        shade(t.cell(i, 0), LIGHT)
        answer = scope.get(key)
        if not fill:
            write(t.cell(i, 1), tr("☐ Yes  ☐ No"))
            write(t.cell(i, 2), hint, hint=True)
            continue
        write(t.cell(i, 1), {True: tr("Yes"), False: tr("No"), None: TBD}[answer], bold=answer is not None)
        detail = scope.get(f"{key}_detail")
        write(t.cell(i, 2), detail or "-", hint=not detail)

    section(doc, tr("Closure reason"))
    reason = d.get("reason") if fill else None
    details = {"REPLACEMENT": tr("name the project that replaces it"), "CANCELLATION": tr("describe the reason")}
    for op in REASONS:
        mark = "☒" if reason == op else "☐"
        text = f"{mark}  {tr(op.capitalize())}"
        if op in details:
            text += f": {d.get('reason_detail')}" if reason == op and d.get("reason_detail") else f" ({details[op]})"
        paragraph(doc, text, bold=reason == op)
    if reason == "COMPLETION" and scope.get("completed") is True:
        paragraph(doc, tr("The deliverables in the project scope are hereby declared completed, and the applicable "
                       "documentation has been prepared, stored and made available to the responsible areas."))

    section(doc, tr("Final remarks"))
    paragraph(doc, v(d.get("remarks")) or " ")

    section(doc, tr("Signatures"))
    signatures = d.get("signatures", {})
    roles = [(tr("Project manager"), "manager"), ("Sponsor", "sponsor"), ("PMO", "pmo")]
    t = table(doc, len(roles) + 1, [4.0, 6.0, 4.5, 2.5])
    header_row(t, [tr("Role"), tr("Name"), tr("Signature"), tr("Date")])
    for i, (role, key) in enumerate(roles, 1):
        write(t.cell(i, 0), role, bold=True)
        write(t.cell(i, 1), v(signatures.get(key)))
        t.rows[i].height = Cm(1.0)

    doc.core_properties.author = doc.core_properties.last_modified_by = "pmo-ai-skills"
    doc.core_properties.title = tr("Project Closure")
    doc.save(target)
    return comp


def validate(d):
    errors = []
    base, actual = d.get("baseline", {}), d.get("actual", {})
    for label, value in (("baseline.start_date", base.get("start_date")), ("baseline.end_date", base.get("end_date")),
                         ("actual.end_date", actual.get("end_date"))):
        if value is not None:
            try:
                date(value)
            except (TypeError, ValueError):
                errors.append(f"{label} must use the DD/MM/YYYY format")
    if not errors and base.get("start_date") and base.get("end_date") and date(base["end_date"]) < date(base["start_date"]):
        errors.append("baseline.end_date cannot be earlier than the start date")
    for label, value in (("baseline.cost", base.get("cost")), ("actual.cost", actual.get("cost"))):
        if value is not None and (type(value) not in (int, float) or value < 0):
            errors.append(f"{label} must be a non-negative number")
    for i, c in enumerate(d.get("approved_changes", []), 1):
        if c.get("schedule_days") is not None and type(c["schedule_days"]) is not int:
            errors.append(f"approved_changes[{i}].schedule_days must be a whole number")
        if c.get("cost") is not None and type(c["cost"]) not in (int, float):
            errors.append(f"approved_changes[{i}].cost must be a number")
    scope = d.get("scope", {})
    for key, required_when in (("completed", False), ("accepted", False), ("open_items", True)):
        value = scope.get(key)
        if value is not None and not isinstance(value, bool):
            errors.append(f"scope.{key} must be true or false")
        elif value is required_when and not scope.get(f"{key}_detail"):
            errors.append(f"scope.{key}_detail is required when scope.{key} is {str(required_when).lower()}")
    reason = d.get("reason")
    if reason is not None and reason not in REASONS:
        errors.append(f"reason must be one of: {', '.join(REASONS)}")
    elif reason in ("REPLACEMENT", "CANCELLATION") and not d.get("reason_detail"):
        errors.append(f"reason_detail is required when the reason is {reason}")
    if errors:
        sys.exit("Invalid input:\n- " + "\n- ".join(errors))


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
    comp = generate(data, args[1], fill=True)
    parts = []
    for label, key, fmt in (("schedule", "schedule", lambda x: f"{x:+d} days"), ("cost", "cost", lambda x: money(x, True))):
        x = comp[key]
        parts.append(f"{label} {fmt(x['deviation'])} ({pct(x['deviation_pct'])})" if x.get("deviation_pct") is not None else f"{label} {TBD}")
    print(f"Closure: {args[1]} | deviation from the updated plan: " + ", ".join(parts))


if __name__ == "__main__":
    main()
