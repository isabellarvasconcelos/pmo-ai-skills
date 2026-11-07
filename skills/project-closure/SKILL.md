---
name: project-closure
description: Generates the Project Closure document as a DOCX file, with lessons learned, planned x actual schedule and cost (accounting for approved changes), scope status, closure reason and signatures. Use when the user asks to formally close a project.
---

# Project Closure

**Project life cycle phase:** closing.

## Workflow

1. Gather the original baseline from the charter: start date, planned end date and total budget.
2. Gather the approved changes (change requests), with the impact of each one on the end date and on cost. They update the baseline. If none are provided, confirm with the user that there were no changes, because the document will compare the actuals with the original baseline.
3. Gather the actuals: actual end date and actual cost.
4. Ask about the scope (completed, accepted, open items) and the closure reason. If there is a lessons learned register, use the most relevant lessons with the same `id` (`LL001`...).
5. Build a JSON file with the structure of `examples/input.json`. Leave out anything that was not provided.
6. Run `python scripts/generate_closure.py input.json output.docx`.
7. Deliver the file and summarize the schedule and cost deviations, separating what came from approved changes from execution deviations.

## Comparison

| Column | Calculation |
|---|---|
| Original plan | Charter baseline |
| Approved changes | Sum of the impacts of approved change requests |
| Updated plan | Original + approved changes |
| Actual | Actual value at closure |
| Deviation | Actual - updated plan, also as % |

Schedule is the duration in calendar days, from the start date to the end date. Comparing with the updated plan avoids treating as a delay or overrun what was formally approved.

## Rules

- **Never make up information.** Anything not provided shows as "TBD", and the deviation is only computed when there is enough data.
- If the scope was not completed or accepted, or if there are open items, a justification is required.
- Use as `reason`: `COMPLETION`, `REPLACEMENT` or `CANCELLATION`. Replacement and cancellation require `reason_detail`.
- The statement that all deliverables were completed only appears when the reason is completion and the scope was completed.
- The document must be signed by the same people who signed the charter.
- Keep a direct, executive tone.

## Language

- The document is generated in English by default. If the user writes in Portuguese or asks for Portuguese, add `"language": "pt-BR"` to the input JSON (or pass `--lang pt-BR` to the script).
- In pt-BR, labels, amounts (R$ 1.234,56) and empty fields ("a definir") come out in Brazilian Portuguese.
- Keep the fixed values listed above as English codes in the input, even in pt-BR. The script writes them translated in the document. Write free text in the document's language.
- Example in Portuguese: `examples/input.pt-BR.json` generates `examples/output.pt-BR.docx`. Blank template: `assets/closure_template.pt-BR.docx`.
