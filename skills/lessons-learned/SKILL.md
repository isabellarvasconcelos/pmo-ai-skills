---
name: lessons-learned
description: Generates and updates a project's Lessons Learned register as an XLSX file, recording positive and negative points, related area, action plan, owner and status. Use when the user asks to record learnings during or at the end of a project.
---

# Lessons Learned

**Project life cycle phase:** monitoring and controlling (continuous logging) and closing (lessons learned session).

## Workflow

1. Gather the lessons from the material provided: accounts, minutes, change requests, risks that materialized or the closing session.
2. Write each lesson as a fact and its consequence ("The API documentation was outdated, which delayed the integration"), not as an opinion or blame.
3. For each lesson, propose an action plan that changes the process for future projects, and confirm the owner with the user.
4. Build a JSON file with the structure of `examples/input.json`. Leave out anything that was not provided.
5. Run `python scripts/generate_lessons.py input.json output.xlsx`.
6. Deliver the file, summarize how many lessons are positive and negative and list the actions still open.

When updating, keep the `id` of existing lessons. New lessons get the next free number.

The project closure reuses these lessons: use the same `id` (`LL001`, `LL002`...) in both documents.

## Rules

- **Never make up lessons, owners or dates.** Anything not provided shows as "TBD".
- Record the positive points too. They show what should be repeated.
- Use as `type`: `POSITIVE` or `NEGATIVE`.
- Use as `status`: `OPEN`, `IN PROGRESS` or `DONE`.
- Use as `timing`: `DURING THE PROJECT`, `AT CLOSURE` or `FUTURE PROJECTS`, depending on when the action plan will be applied.
- Fill in the closing date only for completed actions.
- Keep a direct, executive tone.
