---
name: status-report
description: Generates a project Status Report as a one-slide PPTX, with traffic light, planned x actual %, S-curve, upcoming activities, late activities and risks. Use when the user asks for the periodic status report of a project.
---

# Status Report

**Project life cycle phase:** monitoring and controlling.

## Workflow

1. Gather the data for the period: planned and actual %, S-curve, upcoming activities, late activities and active risks.
2. If there is a risk matrix, use the open items rated "Significant" and "Important".
3. Ask for the project criticality if it was not given, because it defines the traffic light and the reporting frequency.
4. Build a JSON file with the structure of `examples/input.json`. Leave out anything that was not provided.
5. Run `python scripts/generate_status.py input.json output.pptx`.
6. Deliver the file, state the traffic light and the deviation and, if the light is red, say who must be escalated to.

## Traffic light

The deviation is relative to the plan: `(planned - actual) / planned`. A project ahead of plan has a negative deviation.

| Criticality | Frequency | Escalation threshold | Action above the threshold |
|---|---|---|---|
| LOW | Every two weeks | 15% | Escalate to the project manager |
| MEDIUM | Weekly | 10% | Escalate to the sponsor and the manager |
| HIGH | Weekly, with a monthly executive meeting | 5% | Call managers and executive leadership |

- **Green:** deviation up to half the threshold.
- **Yellow:** deviation above half the threshold, up to the threshold.
- **Red:** deviation above the threshold, which requires escalation.
- **Gray:** planned %, actual % or criticality is missing.

## Rules

- **Never make up information.** Anything not provided shows as "TBD".
- Do not choose the light color by hand. The script computes it from the data.
- The slide has limited space. If there are more items than fit, the script says how many were left out. Prioritize the most critical ones and mention the others in your reply.
- For risks, `severity` matches the impact level in the risk matrix.
- Keep a direct, executive tone.

## Language

- The document is generated in English by default. If the user writes in Portuguese or asks for Portuguese, add `"language": "pt-BR"` to the input JSON (or pass `--lang pt-BR` to the script).
- In pt-BR, labels, amounts (R$ 1.234,56) and empty fields ("a definir") come out in Brazilian Portuguese.
- Keep the fixed values listed above as English codes in the input, even in pt-BR. The script writes them translated in the document. Write free text in the document's language.
- Example in Portuguese: `examples/input.pt-BR.json` generates `examples/output.pt-BR.pptx`.
