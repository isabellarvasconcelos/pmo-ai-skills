#!/usr/bin/env python3
"""Generates meeting minutes as DOCX. Usage: generate_minutes.py [input.json] output.docx (no input: blank template)."""
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
STATUS = ["Pending", "In progress", "Done"]
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


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    style_run(p.add_run(text))


def paragraph(doc, text, hint=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    style_run(p.add_run(text), color=HINT if hint else None, italic=hint)


def page_number(footer_paragraph):
    style_run(footer_paragraph.add_run("Page "), color=HINT, size=8)
    run = footer_paragraph.add_run()
    style_run(run, color=HINT, size=8)
    run._r.append(element("fldChar", fldCharType="begin"))
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    run._r.append(instr)
    run._r.append(element("fldChar", fldCharType="separate"))
    result = OxmlElement("w:t")
    result.text = "1"  # placeholder for viewers that do not compute fields; Word updates it
    run._r.append(result)
    run._r.append(element("fldChar", fldCharType="end"))


def generate(d, target, fill):
    def v(value, hint=""):
        if not fill:
            return hint, True
        return (value, False) if value not in (None, "", []) else (TBD, False)

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(2)
    for style in ("Normal", "List Bullet"):
        doc.styles[style].font.name = FONT
        doc.styles[style].font.size = Pt(10)
    page_number(sec.footer.paragraphs[0])

    style_run(doc.add_paragraph().add_run("MEETING MINUTES"), bold=True, color=PRIMARY, size=16)

    ident = d.get("identification", {})
    participants = d.get("participants", [])
    participant_text = "; ".join(f"{p['name']} ({p['role']})" if p.get("role") else p["name"] for p in participants)
    fields = [("Title", ident.get("title")), ("Project or area", ident.get("project")), ("Date", ident.get("date")),
              ("Time", ident.get("time")), ("Duration", ident.get("duration")), ("Location or platform", ident.get("location")),
              ("Participants", participant_text)]
    t = table(doc, len(fields), [4.5, WIDTH - 4.5])
    for i, (label, value) in enumerate(fields):
        write(t.cell(i, 0), label, bold=True)
        shade(t.cell(i, 0), LIGHT)
        text, hint = v(value)
        write(t.cell(i, 1), text, hint=hint)

    section(doc, "Meeting objective")
    paragraph(doc, *v(d.get("objective"), "Why the meeting took place and what it needed to produce."))

    section(doc, "Executive summary")
    summary = d.get("summary")
    if fill and isinstance(summary, list) and summary:
        for item in summary:
            bullet(doc, item)
    else:
        paragraph(doc, *v(summary, "Main points discussed, in a few lines."))

    initiatives = d.get("initiatives", [])
    if initiatives or not fill:
        section(doc, "Mapped initiatives and projects")
        rows = initiatives or [{}] * 2
        t = table(doc, len(rows) + 1, [5.0, WIDTH - 5.0])
        header_row(t, ["Category", "Initiative or project"])
        for i, item in enumerate(rows, 1):
            write(t.cell(i, 0), item.get("category") or (TBD if fill else ""))
            write(t.cell(i, 1), item.get("name") or (TBD if fill else ""))

    section(doc, "Decisions")
    decisions = d.get("decisions", []) if fill else []
    if fill and not decisions:
        paragraph(doc, "No decisions recorded in this meeting.")
    else:
        rows = decisions or [""] * 3
        t = table(doc, len(rows) + 1, [1.2, WIDTH - 1.2])
        header_row(t, ["#", "Decision"])
        for i, decision in enumerate(rows, 1):
            write(t.cell(i, 0), str(i))
            write(t.cell(i, 1), decision)

    section(doc, "Action items")
    actions = d.get("actions", []) if fill else []
    if fill and not actions:
        paragraph(doc, "No action items recorded in this meeting.")
    else:
        rows = actions or [{}] * 3
        t = table(doc, len(rows) + 1, [1.2, 7.8, 3.4, 2.2, 2.4])
        header_row(t, ["#", "Action item", "Owner", "Due date", "Status"])
        for i, item in enumerate(rows, 1):
            write(t.cell(i, 0), str(i))
            for j, key in enumerate(("description", "owner", "due_date", "status"), 1):
                write(t.cell(i, j), item.get(key) or (TBD if fill else ""))

    section(doc, "Next steps")
    steps = d.get("next_steps", [])
    if fill and steps:
        for step in steps:
            bullet(doc, step)
    else:
        paragraph(doc, *v(None, "Next meeting, expected deliverables and follow-ups."))

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    text, _ = v(d.get("prepared_by"), "[name]")
    style_run(p.add_run(f"Minutes prepared by {text}."), color=HINT, italic=True, size=9)

    doc.core_properties.author = doc.core_properties.last_modified_by = "pmo-ai-skills"
    doc.core_properties.title = "Meeting Minutes"
    doc.save(target)


def validate(d):
    errors = []
    for i, p in enumerate(d.get("participants", []), 1):
        if not p.get("name"):
            errors.append(f"participants[{i}].name is required")
    for i, a in enumerate(d.get("actions", []), 1):
        if not a.get("description"):
            errors.append(f"actions[{i}].description is required")
        if a.get("status") is not None and a["status"] not in STATUS:
            errors.append(f"actions[{i}].status must be one of: {', '.join(STATUS)}")
    for i, dec in enumerate(d.get("decisions", []), 1):
        if not isinstance(dec, str) or not dec.strip():
            errors.append(f"decisions[{i}] must be a text")
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
    actions = data.get("actions", [])
    gaps = sum(1 for a in actions if not a.get("owner") or not a.get("due_date"))
    print(f"Minutes: {sys.argv[2]} | {len(data.get('decisions', []))} decision(s) | {len(actions)} action(s)"
          + (f" | {gaps} action(s) without owner or due date" if gaps else ""))


if __name__ == "__main__":
    main()
