# PMO AI Skills

**English** | [Português](README.pt-BR.md)

AI skills that turn project management steps into deliverables ready for review.

This repository comes from real PMO practice, rewritten in a generic version. All examples use fictitious data.

## Skills

| Skill | Phase | Output |
|---|---|---|
| [Grill Me: Projects](skills/grill-me-projects/SKILL.md) | Before initiation | Discovery interview that ends in a PRD |
| [Project Charter](skills/project-charter/SKILL.md) | Initiation | XLSX spreadsheet with the charter and procurement |
| [Risk, Issue and Opportunity Matrix](skills/risk-issue-opportunity-matrix/SKILL.md) | Planning, monitoring and controlling | XLSX spreadsheet with the item log and computed rating |
| [Status Report](skills/status-report/SKILL.md) | Monitoring and controlling | PPTX slide with traffic light, S-curve, activities and risks |
| [Change Request](skills/change-request/SKILL.md) | Monitoring and controlling | DOCX document with the change, impact analysis and approval |
| [Meeting Minutes](skills/meeting-minutes/SKILL.md) | All | DOCX document with decisions and action items, including from loose notes |
| [Project Closure](skills/project-closure/SKILL.md) | Closing | DOCX document with lessons, planned x actual and scope status |
| [Lessons Learned](skills/lessons-learned/SKILL.md) | Monitoring and controlling, closing | XLSX spreadsheet with positive and negative points and action plans |

## How to try it

Requires Python 3.9 or later.

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

cd skills/project-charter
python scripts/generate_charter.py examples/input.json my_charter.xlsx
```

Skills that generate files have their own script, plus a sample input and the generated output in `examples/`. Grill Me is a conversation only: see the [sample session](skills/grill-me-projects/examples/discovery-session.md), also available [in Portuguese](skills/grill-me-projects/examples/discovery-session.pt-BR.md).

All skills use the same fictitious project, a digital delivery check-in, so the numbers match from one document to the next.

## Languages

Documents are generated in English by default. To generate them in Brazilian Portuguese, add `"language": "pt-BR"` to the input or pass `--lang pt-BR`:

```bash
python scripts/generate_charter.py examples/input.pt-BR.json my_charter.xlsx
python scripts/generate_charter.py --lang pt-BR blank_charter.xlsx
```

Each skill has a Portuguese example (`input.pt-BR.json` and `output.pt-BR.*`), and the blank templates exist in both languages.

## Principles

- **Never make up data:** anything not provided shows as "TBD".
- **Keep people in control:** the AI prepares the document, and review and approval remain human.
- **Validate what can be validated:** invalid values and types are rejected, and totals come from formulas.

## Credits

The [Grill Me: Projects](skills/grill-me-projects/SKILL.md) skill adapts Matt Pocock's [grill-me](https://github.com/mattpocock/skills/tree/main/skills/productivity/grill-me) (MIT), based on a version by Tiago Ornelas. See the [full credits](skills/grill-me-projects/CREDITS.md).

## License

[MIT](LICENSE). You may use, adapt and redistribute it, keeping the copyright notice. Third-party adapted parts also keep their original license, as stated in the credits.
