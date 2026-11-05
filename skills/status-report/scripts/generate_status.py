#!/usr/bin/env python3
"""Generates the status report as a one-slide PPTX. Usage: generate_status.py input.json output.pptx"""
import json
import sys
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

TBD = "TBD"
FONT = "Arial"
PRIMARY, LIGHT, BORDER, TEXT, MUTED = "1F3A5F", "DCE6F1", "BFBFBF", "262626", "7F7F7F"
CRITICALITY = {"LOW": 15, "MEDIUM": 10, "HIGH": 5}  # escalation threshold, in %
ACTIONS = {"LOW": "escalate to the project manager", "MEDIUM": "escalate to the sponsor and the manager",
           "HIGH": "call managers and executive leadership"}
LIGHTS = {"Green": ("2E8B57", "On track"), "Yellow": ("E0A800", "Attention"),
          "Red": ("C0392B", "At risk of delay"), "Gray": ("9E9E9E", TBD)}
LIMITS = {"upcoming_activities": 5, "late_activities": 5, "risks": 4}


def color(hex_code):
    return RGBColor.from_string(hex_code)


def is_number(v):
    return type(v) in (int, float)


def validate(d):
    errors = []
    for field in ("planned", "actual"):
        v = d.get(field)
        if v is not None and (not is_number(v) or not 0 <= v <= 100):
            errors.append(f"{field} must be a number from 0 to 100")
    if d.get("criticality") is not None and d["criticality"] not in CRITICALITY:
        errors.append(f"criticality must be one of: {', '.join(CRITICALITY)}")
    curve = d.get("s_curve")
    if curve:
        n = len(curve.get("periods", []))
        for series in ("planned", "actual"):
            values = curve.get(series, [])
            if len(values) > n:
                errors.append(f"s_curve.{series} has more values than periods")
            filled = [v for v in values if v is not None]
            if any(not is_number(v) or not 0 <= v <= 100 for v in filled):
                errors.append(f"s_curve.{series} must contain numbers from 0 to 100")
            elif filled != sorted(filled):
                errors.append(f"s_curve.{series} must be cumulative (it cannot decrease)")
    for i, r in enumerate(d.get("risks", []), 1):
        for field in ("probability", "severity"):
            v = r.get(field)
            if v is not None and (type(v) is not int or not 1 <= v <= 5):
                errors.append(f"risks[{i}].{field} must be an integer from 1 to 5")
    if errors:
        sys.exit("Invalid input:\n- " + "\n- ".join(errors))


def traffic_light(d):
    planned, actual, crit = d.get("planned"), d.get("actual"), d.get("criticality")
    if planned is None or actual is None or crit is None:
        return "Gray", None
    if planned == 0:
        return "Green", 0.0
    deviation = (planned - actual) / planned * 100
    threshold = CRITICALITY[crit]
    if deviation > threshold:
        return "Red", deviation
    return ("Yellow" if deviation > threshold / 2 else "Green"), deviation


def text(slide, x, y, w, h, value, size=9, bold=False, text_color=TEXT, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.margin_left = tf.margin_right = Inches(0.03)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = str(value)
    r.font.name, r.font.size, r.font.bold = FONT, Pt(size), bold
    r.font.color.rgb = color(text_color)
    return box


def rectangle(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = color(fill)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = color(line)
        s.line.width = Pt(0.75)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s


def panel(slide, x, y, w, h, title):
    rectangle(slide, x, y, w, h, line=BORDER)
    text(slide, x + 0.1, y + 0.02, w - 0.2, 0.26, title, size=11, bold=True, text_color=PRIMARY)
    rectangle(slide, x + 0.1, y + 0.3, w - 0.2, 0.012, fill=PRIMARY)


def table(slide, x, y, w, columns, rows, list_key, warnings):
    """columns: (title, width fraction, function that extracts the value)."""
    limit = LIMITS[list_key]
    if len(rows) > limit:
        warnings.append(f"{list_key}: {len(rows) - limit} item(s) did not fit on the slide")
    rows = rows[:limit]
    if not rows:
        text(slide, x, y, w, 0.25, TBD, text_color=MUTED)
        return
    shape = slide.shapes.add_table(len(rows) + 1, len(columns), Inches(x), Inches(y), Inches(w), Inches(0.22 * (len(rows) + 1)))
    t = shape.table
    t.first_row = True
    for j, (_, fraction, _) in enumerate(columns):
        t.columns[j].width = Inches(w * fraction)
    for i in range(len(rows) + 1):
        t.rows[i].height = Inches(0.22)
        for j, (title, _, extract) in enumerate(columns):
            c = t.cell(i, j)
            value = title if i == 0 else extract(rows[i - 1])
            c.text = TBD if value in (None, "") else str(value)
            c.margin_left = c.margin_right = Inches(0.04)
            c.margin_top = c.margin_bottom = Inches(0.01)
            c.fill.solid()
            c.fill.fore_color.rgb = color(PRIMARY if i == 0 else ("FFFFFF" if i % 2 else "F5F8FC"))
            for p in c.text_frame.paragraphs:
                for r in p.runs:
                    r.font.name, r.font.size, r.font.bold = FONT, Pt(8), i == 0
                    r.font.color.rgb = color("FFFFFF" if i == 0 else TEXT)


def s_curve(slide, x, y, w, h, curve):
    if not curve or not curve.get("periods"):
        text(slide, x, y, w, 0.25, TBD, text_color=MUTED)
        return
    data = CategoryChartData(number_format='0"%"')
    data.categories = curve["periods"]
    n = len(curve["periods"])
    for name, series in (("Planned", "planned"), ("Actual", "actual")):
        values = list(curve.get(series, []))
        data.add_series(name, values + [None] * (n - len(values)))
    chart = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(x), Inches(y), Inches(w), Inches(h), data).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.TOP
    chart.legend.include_in_layout = False
    chart.legend.font.size, chart.legend.font.name = Pt(8), FONT
    for axis in (chart.category_axis, chart.value_axis):
        axis.tick_labels.font.size, axis.tick_labels.font.name = Pt(8), FONT
        axis.format.line.color.rgb = color(BORDER)
    chart.value_axis.maximum_scale, chart.value_axis.minimum_scale = 100, 0
    chart.value_axis.major_unit = 25
    chart.value_axis.major_gridlines.format.line.color.rgb = color("E7E7E7")
    for series, hex_code, dashed in zip(chart.plots[0].series, (MUTED, PRIMARY), (True, False)):
        series.smooth = False
        series.format.line.color.rgb = color(hex_code)
        series.format.line.width = Pt(2)
        if dashed:
            series.format.line.dash_style = 4  # dash
        series.marker.format.fill.solid()
        series.marker.format.fill.fore_color.rgb = color(hex_code)
        series.marker.format.line.color.rgb = color(hex_code)


def generate(d, target):
    ident = d.get("identification", {})
    v = lambda key: ident.get(key) or TBD
    pct = lambda key: f"{d[key]:g}%" if d.get(key) is not None else TBD
    light, deviation = traffic_light(d)
    warnings = []

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    s = prs.slides.add_slide(prs.slide_layouts[6])

    rectangle(s, 0, 0, 10, 0.62, fill=PRIMARY)
    text(s, 0.2, 0.06, 6.3, 0.5, f"STATUS REPORT | {v('title')}", size=15, bold=True, text_color="FFFFFF")
    for i, (label, value) in enumerate((("Code:", v("code")), ("Status date:", v("status_date")))):
        text(s, 6.6, 0.11 + i * 0.2, 1.3, 0.18, label, size=9, text_color="FFFFFF", align=PP_ALIGN.RIGHT)
        text(s, 7.95, 0.11 + i * 0.2, 1.95, 0.18, value, size=9, bold=True, text_color="FFFFFF")

    rectangle(s, 0, 0.62, 10, 0.45, fill=LIGHT)
    for i, (label, value) in enumerate((("Start:", v("start_date")), ("End:", v("end_date")))):
        text(s, 0.2, 0.66 + i * 0.18, 0.55, 0.18, label, bold=True)
        text(s, 0.75, 0.66 + i * 0.18, 1.0, 0.18, value)
    text(s, 1.8, 0.66, 0.85, 0.36, "Criticality:", bold=True)
    text(s, 2.62, 0.66, 0.8, 0.36, d.get("criticality") or TBD)
    for x, label, value in ((3.5, "Planned:", pct("planned")), (4.85, "Actual:", pct("actual"))):
        text(s, x, 0.66, 0.75, 0.36, label, bold=True)
        text(s, x + 0.75, 0.66, 0.6, 0.36, value, size=12 if value != TBD else 9, bold=True, text_color=PRIMARY)
    text(s, 6.2, 0.66, 0.75, 0.36, "Deviation:", bold=True)
    text(s, 6.9, 0.66, 0.7, 0.36, f"{deviation:.1f}%" if deviation is not None else TBD,
         size=12 if deviation is not None else 9, bold=True, text_color=PRIMARY)
    hex_code, label = LIGHTS[light]
    text(s, 7.65, 0.66, 0.55, 0.36, "Status:", bold=True)
    rectangle(s, 8.2, 0.73, 0.22, 0.22, fill=hex_code, shape=MSO_SHAPE.OVAL)
    text(s, 8.48, 0.66, 1.45, 0.36, label, bold=True, text_color=hex_code)

    panel(s, 0.2, 1.15, 4.75, 2.05, "Project S-curve")
    s_curve(s, 0.3, 1.48, 4.55, 1.68, d.get("s_curve"))

    panel(s, 5.05, 1.15, 4.75, 2.05, "Upcoming activities")
    table(s, 5.15, 1.55, 4.45, [("Activity", 0.55, lambda a: a.get("activity")), ("Owner", 0.25, lambda a: a.get("owner")),
                                ("End date", 0.2, lambda a: a.get("end_date"))], d.get("upcoming_activities", []), "upcoming_activities", warnings)

    panel(s, 0.2, 3.3, 4.75, 2.15, "Late or at-risk activities")
    table(s, 0.3, 3.7, 4.45, [("Activity", 0.55, lambda a: a.get("activity")), ("Area", 0.25, lambda a: a.get("area")),
                              ("End date", 0.2, lambda a: a.get("end_date"))], d.get("late_activities", []), "late_activities", warnings)

    panel(s, 5.05, 3.3, 4.75, 2.15, "Project risks")
    table(s, 5.15, 3.7, 4.45, [("Risk", 0.38, lambda r: r.get("description")), ("Prob.", 0.12, lambda r: r.get("probability")),
                               ("Sev.", 0.11, lambda r: r.get("severity")), ("Action plan", 0.39, lambda r: r.get("action_plan"))],
          d.get("risks", []), "risks", warnings)

    prs.core_properties.author = prs.core_properties.last_modified_by = "pmo-ai-skills"
    prs.core_properties.title = "Status Report"
    prs.save(target)
    return light, deviation, warnings


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    data = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    validate(data)
    light, deviation, warnings = generate(data, sys.argv[2])
    summary = f"Status: {sys.argv[2]} | light {light}"
    if deviation is not None:
        summary += f" | deviation {deviation:.1f}%"
    if light == "Red":
        summary += f" | escalate: {ACTIONS[data['criticality']]}"
    print(summary)
    for warning in warnings:
        print(f"Warning: {warning}")


if __name__ == "__main__":
    main()
