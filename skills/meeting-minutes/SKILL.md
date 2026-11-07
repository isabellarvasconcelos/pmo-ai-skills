---
name: meeting-minutes
description: Generates Meeting Minutes as a DOCX file, with identification, objective, executive summary, mapped initiatives, decisions, action items and next steps. Use when the user asks to produce, write or formalize the minutes of a meeting, alignment, committee or project follow-up, including from transcripts or notes.
---

# Meeting Minutes

**Project life cycle phase:** cross-cutting, in any phase where a meeting needs to be formalized.

## Workflow

1. Read the material provided: transcript, notes or account of the meeting.
2. Split the content into the sections of the minutes. Tell what was **decided** apart from what was only **discussed**: only what was settled in the meeting goes under "Decisions".
3. For each action item, identify the owner and the due date. If the material does not say, use "TBD" and point out the gap in your reply.
4. Build a JSON file with the structure of `examples/input.json`. Leave out anything that was not provided.
5. Run `python scripts/generate_minutes.py input.json output.docx`.
6. Deliver the file and list the gaps: items without an owner or due date and points that looked decided but remained ambiguous.

See `examples/meeting-notes.md` for an example of loose notes that became the minutes in `examples/output.docx`.

## Rules

- **Never make up participants, decisions or action items** that are not in the material.
- Record decisions as affirmative, verifiable sentences ("CR-001 approved"), not as topics ("Discussion about Wi-Fi").
- Use as action item `status`: `Pending`, `In progress` or `Done`.
- The mapped initiatives section only appears when there are initiatives. It records new demands raised in the meeting, for later prioritization.
- Keep a direct, executive tone.

## Language

- The document is generated in English by default. If the user writes in Portuguese or asks for Portuguese, add `"language": "pt-BR"` to the input JSON (or pass `--lang pt-BR` to the script).
- In pt-BR, labels, amounts (R$ 1.234,56) and empty fields ("a definir") come out in Brazilian Portuguese.
- Keep the fixed values listed above as English codes in the input, even in pt-BR. The script writes them translated in the document. Write free text in the document's language.
- Example in Portuguese: `examples/input.pt-BR.json` generates `examples/output.pt-BR.docx`. Blank template: `assets/minutes_template.pt-BR.docx`. Loose notes in Portuguese: `examples/meeting-notes.pt-BR.md`.
