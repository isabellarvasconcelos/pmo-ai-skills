---
name: project-charter
description: Generates the Project Charter as an XLSX file, with business case, SMART objective, AS IS x TO BE, scope, requirements, technical solution, impacts, resources, budget and procurement. Use when the user asks to formally open a project or document its initial scope.
---

# Project Charter

**Project life cycle phase:** initiation.

## Workflow

1. Gather what the user has already provided, without asking again about what is already clear.
2. Ask only for what is missing to understand the business case, objective, scope and budget. If the user does not know, move on.
3. Build a JSON file with the structure of `examples/input.json`. Leave out anything that was not provided.
4. Run `python scripts/generate_charter.py input.json output.xlsx`.
5. Deliver the file and list the fields left as "TBD" and the scope points that deserve review.

## Rules

- **Never make up information.** Fields with no data are filled with "TBD", which the script does automatically.
- Write the objective in SMART form: action verb, measurable result and deadline. If the metric or deadline is missing, point out the gap instead of creating it.
- State exclusions under "Out of scope", because they prevent scope disputes later.
- Treat every assumption as a possible risk and suggest adding it to the risk matrix.
- Do not edit totals by hand. The budget uses formulas, and the total of the "Procurement" sheet feeds the "Charter" sheet.
- Use only `CAPEX` or `OPEX` as `type`. The script rejects other values and invalid numbers.
- Keep a direct, executive tone.
