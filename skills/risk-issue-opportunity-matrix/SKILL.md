---
name: risk-issue-opportunity-matrix
description: Generates and updates a project's Risk, Issue and Opportunity Matrix as an XLSX file, computing each item's rating from impact x probability. Use when the user asks to log, update or review risks, issues or opportunities of a project.
---

# Risk, Issue and Opportunity Matrix

**Project life cycle phase:** planning, monitoring and controlling (continuous review).

## Workflow

1. Gather the items the user provided. If there is a project charter, read its assumptions and constraints, because each assumption may create a risk.
2. Suggest items that seem to be missing, but only add them after the user confirms.
3. Ask the user for the impact and probability of each item, or propose scores for the user to confirm.
4. Build a JSON file with the structure of `examples/input.json`. Leave out anything that was not provided.
5. Run `python scripts/generate_matrix.py input.json output.xlsx`.
6. Deliver the file, highlight the "Significant" items and list what was left as "TBD".

When updating, keep the `id` of existing items. New items get the next free number.

## Rating

The rating is impact (1 to 5) x probability (1 to 5):

| Score | Rating |
|---|---|
| 1 to 4 | Not significant |
| 5 to 12 | Important |
| 15 to 25 | Significant |

## Rules

- **Never make up information.** Scores, costs and owners with no data are filled with "TBD".
- Tell the types apart: a `RISK` may happen, an `ISSUE` has already happened and an `OPPORTUNITY` is a possible positive effect.
- Use as `response`: `Mitigate`, `Share`, `Transfer`, `Avoid` or `Accept` for risks and issues; `Exploit`, `Share`, `Enhance` or `Accept` for opportunities.
- Use as `status`: `OPEN`, `CLOSED` or `CANCELLED`.
- Do not type the rating by hand. It comes from a formula in the spreadsheet.
- Keep a direct, executive tone.

## Language

- The document is generated in English by default. If the user writes in Portuguese or asks for Portuguese, add `"language": "pt-BR"` to the input JSON (or pass `--lang pt-BR` to the script).
- In pt-BR, labels, amounts (R$ 1.234,56) and empty fields ("a definir") come out in Brazilian Portuguese.
- Keep the fixed values listed above as English codes in the input, even in pt-BR. The script writes them translated in the document. Write free text in the document's language.
- Example in Portuguese: `examples/input.pt-BR.json` generates `examples/output.pt-BR.xlsx`. Blank template: `assets/matrix_template.pt-BR.xlsx`. New IDs use the `RISC` prefix in pt-BR (`RISC001`).
