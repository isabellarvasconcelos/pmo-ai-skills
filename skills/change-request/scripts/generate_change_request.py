#!/usr/bin/env python3
"""Generates the Change Request as DOCX. Usage: generate_change_request.py [input.json] output.docx (no input: blank template)."""
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

TBD = "TBD"
FONT = "Arial"
PRIMARY, LIGHT, HINT = "1F3A5F", "DCE6F1", "7F7F7F"
ASPECTS = ["Schedule", "Requirements", "Resources", "Costs", "Scope", "Technical", "Functional"]
DECISIONS = ["APPROVED", "REJECTED"]
WIDTH = 17.0  # usable cm on A4 with 2 cm margins
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW",
               "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"]


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


def write(cell, text, bold=False, color=None, hint=False, size=10):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    style_run(p.add_run(str(text)), bold, HINT if hint else color, italic=hint, size=size)


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


def title_row(t, text):
    row = t.rows[0].cells
    cell = row[0].merge(row[-1]) if len(row) > 1 else row[0]
    write(cell, text, bold=True, color="FFFFFF")
    shade(cell, PRIMARY)


def header_row(t, row, titles):
    for j, title in enumerate(titles):
        write(t.cell(row, j), title, bold=True)
        shade(t.cell(row, j), LIGHT)


def spacer(doc):
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def generate(d, target, fill):
    def value(v, hint=""):
        """Returns (text, is_hint). In the blank template, shows the hint."""
        if not fill:
            return hint, True
        return (v, False) if v not in (None, "", []) else (TBD, False)

    def text_block(title, v, hint):
        t = table(doc, 2, [WIDTH])
        title_row(t, title)
        text, is_hint = value(v, hint)
        write(t.cell(1, 0), text, hint=is_hint)
        spacer(doc)

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(2)
    doc.styles["Normal"].font.name, doc.styles["Normal"].font.size = FONT, Pt(10)

    style_run(doc.add_paragraph().add_run("CHANGE REQUEST"), bold=True, color=PRIMARY, size=16)

    ident = d.get("identification", {})
    fields = [("Project code", "code"), ("Project name", "title"), ("Change code", "change_code"),
              ("Requested by", "requested_by"), ("Request date", "request_date")]
    t = table(doc, len(fields), [5.0, WIDTH - 5.0])
    for i, (label, key) in enumerate(fields):
        write(t.cell(i, 0), label, bold=True)
        shade(t.cell(i, 0), LIGHT)
        text, is_hint = value(ident.get(key))
        write(t.cell(i, 1), text, hint=is_hint)
    spacer(doc)

    style_run(doc.add_paragraph().add_run(
        "This document formalizes the project change request and the assessment of the impacts it will cause."), italic=True)

    text_block("Change description", d.get("description"), "What changes, in detail, and in which part of the project.")
    text_block("Justification", d.get("justification"), "Why the change is needed and what happens if it is not made.")

    impacts = {i["aspect"]: i for i in d.get("impacts", [])}
    t = table(doc, len(ASPECTS) + 2, [3.0, 2.2, 7.8, 4.0])
    title_row(t, "Change impact analysis")
    header_row(t, 1, ["Aspect", "Impacted?", "Impact justification", "Reference document"])
    for i, aspect in enumerate(ASPECTS, 2):
        item = impacts.get(aspect)
        write(t.cell(i, 0), aspect, bold=True)
        if not fill:
            write(t.cell(i, 1), "☐ Yes\n☐ No")
            continue
        write(t.cell(i, 1), "Yes" if item else "No", bold=bool(item))
        write(t.cell(i, 2), item["justification"] if item else "No impact identified.", hint=not item)
        write(t.cell(i, 3), (item or {}).get("document") or (TBD if item else "-"))
    spacer(doc)

    summary = d.get("impact_summary", {})
    t = table(doc, 2, [WIDTH / 2] * 2)
    for j, (label, key, fmt) in enumerate((("Impact on project end date", "schedule_days", lambda v: f"{v:+d} calendar days" if v else "End date unchanged"),
                                           ("Cost impact", "cost", lambda v: f"{'+' if v >= 0 else '-'}${abs(v):,.2f}"))):
        write(t.cell(0, j), label, bold=True)
        shade(t.cell(0, j), LIGHT)
        v = summary.get(key)
        text, is_hint = value(fmt(v) if v is not None else None)
        write(t.cell(1, j), text, hint=is_hint, bold=not is_hint and v is not None, size=12)
    spacer(doc)

    risks = d.get("risks_and_impacts", {})
    t = table(doc, 3, [5.0, WIDTH - 5.0])
    title_row(t, "Risks and impacts")
    for i, (label, key, hint) in enumerate((("Risks of the change", "risks", "What could go wrong when carrying out the change."),
                                            ("Impact of not changing", "impacts", "What happens if the change is rejected.")), 1):
        write(t.cell(i, 0), label, bold=True)
        shade(t.cell(i, 0), LIGHT)
        text, is_hint = value(risks.get(key), hint)
        write(t.cell(i, 1), text, hint=is_hint)
    spacer(doc)

    plan = d.get("execution_plan", []) if fill else []
    rows = max(len(plan), 1 if fill else 5)
    t = table(doc, rows + 2, [1.6, 7.0, 3.4, 2.4, 2.6])
    title_row(t, "Change execution plan")
    header_row(t, 1, ["Seq.", "Activity", "Owner", "Date", "Expected duration"])
    for i in range(rows):
        item = plan[i] if i < len(plan) else {}
        write(t.cell(i + 2, 0), str(i + 1))
        for j, key in enumerate(("activity", "owner", "date", "duration"), 1):
            write(t.cell(i + 2, j), item.get(key) or (TBD if fill else ""))
    spacer(doc)

    text_block("Final notes", d.get("notes"), "")

    t = table(doc, 2, [WIDTH])
    title_row(t, "Final decision")
    decision = d.get("decision")
    marks = "     ".join(f"{'☒' if decision == op else '☐'} {op}" for op in DECISIONS)
    if fill and not decision:
        marks += "     (pending review)"
    write(t.cell(1, 0), marks, bold=True)
    spacer(doc)

    roles = ["Requester", "Project manager", "Sponsor"]
    t = table(doc, len(roles) + 1, [4.5, 8.0, 4.5])
    header_row(t, 0, ["Role", "Name and signature", "Date"])
    for i, role in enumerate(roles, 1):
        write(t.cell(i, 0), role, bold=True)
        t.rows[i].height = Cm(1.0)

    doc.core_properties.author = doc.core_properties.last_modified_by = "pmo-ai-skills"
    doc.core_properties.title = "Change Request"
    doc.save(target)


def validate(d):
    errors, seen = [], set()
    for i, item in enumerate(d.get("impacts", []), 1):
        if item.get("aspect") not in ASPECTS:
            errors.append(f"impacts[{i}].aspect must be one of: {', '.join(ASPECTS)}")
        elif item["aspect"] in seen:
            errors.append(f"impacts[{i}].aspect is duplicated: {item['aspect']}")
        seen.add(item.get("aspect"))
        if not item.get("justification"):
            errors.append(f"impacts[{i}].justification is required for an impacted aspect")
    summary = d.get("impact_summary", {})
    if summary.get("schedule_days") is not None and type(summary["schedule_days"]) is not int:
        errors.append("impact_summary.schedule_days must be a whole number of days")
    if summary.get("cost") is not None and type(summary["cost"]) not in (int, float):
        errors.append("impact_summary.cost must be a number")
    if d.get("decision") is not None and d["decision"] not in DECISIONS:
        errors.append(f"decision must be one of: {', '.join(DECISIONS)}")
    if errors:
        sys.exit("Invalid input:\n- " + "\n- ".join(errors))


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
    aspects = ", ".join(i["aspect"] for i in data.get("impacts", [])) or "none"
    print(f"Change request: {sys.argv[2]} | impacted aspects: {aspects} | decision: {data.get('decision') or 'pending'}")


if __name__ == "__main__":
    main()
