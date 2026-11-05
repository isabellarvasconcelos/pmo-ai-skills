---
name: change-request
description: Generates the Change Request form as a DOCX file, with change description, impact analysis by aspect, risks, execution plan and final decision. Use when the user asks to formalize a change in scope, schedule, cost or requirements of a project already in progress.
---

# Change Request

**Project life cycle phase:** monitoring and controlling, at any point during execution.

## Workflow

1. Understand the change: what changes, why it changes and who requested it.
2. Assess the impact on each aspect: schedule, requirements, resources, costs, scope, technical and functional. Ask about anything unclear, and do not mark an aspect as impacted without a justification.
3. If the project has a charter, risk matrix or status report, use them to compare against the baseline and point out what needs to be updated.
4. Build a JSON file with the structure of `examples/input.json`. Leave out anything that was not provided.
5. Run `python scripts/generate_change_request.py input.json output.docx`.
6. Deliver the file, summarize the schedule and cost impact and list the documents that must be updated if the change is approved.

In `impact_summary`, `schedule_days` is the shift of the **project end date**, and `cost` is the budget variation. Delays in activities that do not move the end date belong in the justification of the "Schedule" aspect.

## Rules

- **Never make up information.** Anything not provided shows as "TBD".
- Every aspect marked as impacted needs a justification. The script rejects an impact without one.
- Fill in the decision (`APPROVED` or `REJECTED`) only when the user provides it. Without it, the form is marked as pending review.
- An approved change alters the baseline: remind the user to update the charter, the schedule and the risk matrix.
- Keep a direct, executive tone.
