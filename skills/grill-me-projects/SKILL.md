---
name: grill-me-projects
description: Rigorous discovery interview that builds a shared understanding of a project demand (what the person knows and what existing documents and processes reveal), keeping each session to a single workstream and ending in a PRD ready for validation.
disable-model-invocation: true
---

# Grill Me: Projects

> Adapted from Matt Pocock's grill-me skill (MIT), based on a version by Tiago Ornelas. See [CREDITS.md](CREDITS.md).

**Before you start:** if what the person described is a small, contained change with no relevant impact on process, responsibility, approval, audience or project outcome, say so and offer a short alignment conversation as the faster path.

This is a recommendation, not a barrier: if the person still wants the full interview, proceed normally.

This is a discovery interview. Its outputs are:

1. A shared understanding of the demand.
2. User stories or structured needs.
3. A complete PRD, presented in the conversation.

Keep the session to a single, concise and deep workstream. Do not define the technical solution, architecture, code, database, APIs, development plan or breakdown of engineering tasks. Those decisions belong to the people responsible for technical execution.

## Close both directions of the unknown

Discovery fails when either side holds back relevant knowledge:

- **What the person knows and you do not:** business reason, undocumented constraints, previous attempts, implicit rules, people affected and what a finished delivery really means. Only the interview reveals this.
- **What existing materials know and the person may not:** documented processes, commitments made, policies, previous decisions, dependencies, deadlines, available data and how an operation really works. Research these materials and present the findings.

Do both. An interview that only extracts information from the person is incomplete.

## How to interview

Interview the person rigorously until you reach a shared understanding. Organize it as a **decision tree**: each decision opens the decisions that depend on it.

Work the tree in **rounds**. The **frontier** is the set of decisions whose prerequisites are already settled: the questions that can be asked now, without assuming answers not yet given.

In each round:

- Ask the whole available frontier.
- Number each question.
- Include a reasoned recommendation.
- Wait for the answers before the next round.

Use this format:

```text
❓ **Q1. <question title>**: <question body, including context and alternatives when needed>

➡️ <your recommended answer>
```

Answers reshape the tree. Confirmed decisions open new questions. Recompute the frontier after each round.

A question whose answer depends on another decision still open in the same round belongs to a later round.

### Look up facts; ask for decisions

Finding facts is your job. Do not ask the person for anything that can be obtained from documents, project materials, available tools or the context provided.

When a question depends on a fact available in the environment:

- Look up the fact.
- Present the evidence concisely.
- Continue with the other independent questions.
- Leave only the decisions that depend on that fact for a later round.

When the fact is outside the available materials (for example, a regulation, external policy, contract, vendor behavior or public documentation), research a reliable source before treating it as fact.

Decisions belong to the user. Present options, consequences and a recommendation; wait for the decision.

### When the question needs visual or experiential validation

If a decision is essentially about clarity, appearance, user experience or understanding of a flow, do not drag out an abstract debate.

Propose a simple representation, such as:

- Flowchart.
- Journey.
- Wireframe.
- Sample communication.
- Scenario simulation.

Once the person chooses a direction, treat the choice as a confirmed decision and add it to the stories or needs.

## Protect the workstream scope

Before consolidating the stories, check whether the session is designing more than one independent workstream.

Apply this test in every round:

- Does this decision still belong to the objective, process, outcome and responsibilities of the workstream being defined?
- Or does it start defining another process, another deliverable, another owning area or another initiative with its own outcome?

When you identify two independent workstreams:

1. Name the workstreams that are forming.
2. Propose limiting this session to the most defined one.
3. Record the other as a named future initiative.
4. Add it to "Out of scope" in the PRD, without turning it into a vague pending item.

Do not split a workstream just because it has many decisions. Depth and volume are not, by themselves, scope creep. The trigger is a second independent deliverable, with its own objective, owner or outcome.

### The smaller leak: a single question that belongs to another workstream

Scope creep can appear inside a single, seemingly reasonable question.

Before asking, assess:

> Does the answer decide the outcome, rule or responsibility of this workstream? Or does it decide how another area, process or initiative will react to that outcome?

When a question mixes both:

- Separate the decisions.
- Keep in the interview only the part that belongs to this workstream.
- Record the decision that belongs to the other workstream under "Out of scope", naming who should handle it later.

Example:

- **In scope:** what information or outcome this workstream must provide when there is an exception.
- **Out of scope:** how another area should change its process when it receives that outcome.

## Protect the delivery size

Even within a single workstream, a demand may hold more work than fits in one useful delivery.

As stories or needs take shape, check whether each set can be:

- Understood and validated as a unit.
- Delivered in one step with an observable outcome.
- Useful on its own, without requiring every future step to be finished too.

When the demand is too large:

1. Name the smallest complete delivery that generates value first.
2. Explain what comes next and why.
3. Record the later steps under "Out of scope" as named initiatives.
4. Do not discard the postponed steps.

Do not fragment a delivery that only works as a whole. The goal is the smallest complete unit of value, not the shortest possible list of activities.

## Consolidate into stories or needs

Discovery ends when the work can be stated as a list of stories or needs, in the format:

> As <affected person or area>, I need <capability, information or change>, so that <outcome or benefit>.

When stories are not the best representation, use:

> For <project objective>, <outcome or condition> is needed, so that <benefit or impact>.

Present the complete list and ask for confirmation.

Do not generate the PRD until the person confirms there is a shared understanding.

## Generate the PRD

After confirmation, produce a complete PRD.

The PRD is a project and business document. Describe the problem, the audience, the value, the scope, the expected process and what defines success.

Do not include architecture, implementation, code, technical interfaces, development estimates or a technical execution plan.

### Self-containment rule

Assume the reader does not know:

- The conversation.
- The interview.
- The project.
- The documents consulted.
- The previous context.

Everything needed to understand the demand, its limits and its decisions must be in the PRD.

Never write "as discussed", "as per our previous conversation" or similar.

### Language

Write the PRD in English, unless the person explicitly asks for another language. Keep the same language throughout the document.

### Synthesis process

1. Use the confirmed understanding and the materials consulted to ground the document.
2. Check again whether the demand represents more than one independent delivery.
3. If it does, state the cut of this PRD and name the later deliveries under "Out of scope".
4. Write the PRD using the template below.
5. Present it in full.
6. Follow the delivery and next step process.

<prd-template>

# [Initiative name]

## Problem statement

Describe the problem from the perspective of the people, areas or processes affected.

Include:

- Current context.
- Impact of the problem.
- Available evidence.
- Consequence of not acting.

## Objective and expected outcome

Describe:

- The change the initiative seeks to produce.
- The expected value.
- How it will be possible to recognize that the outcome was achieved.

## People and areas affected

| Group | Role or impact | Need or expectation |
|---|---|---|

## Scope

### Included

List what this initiative covers.

### Out of scope

List what this initiative deliberately does not cover.

When there is related future work, name it and explain why it is outside this cut.

## Needs and stories

A long, numbered list. Use one of the formats below.

1. As <affected person or area>, I need <capability, information or change>, so that <outcome or benefit>.

or:

1. For <project objective>, <outcome or condition> is needed, so that <benefit or impact>.

Under each item, describe its definition of done as an observable behavior or outcome.

The definition of done must state what is true from the business, process or user point of view, never how it will be implemented.

## Expected process or journey

Describe the future flow in business language.

1. Trigger that starts the process.
2. Main steps and decisions.
3. Information, approvals or records needed.
4. Handling of relevant exceptions.
5. Expected final outcome.

## Rules and acceptance criteria

| ID | Condition | Expected outcome | Exceptions or notes |
|---|---|---|---|

Use verifiable criteria, for example:

- Given <context>, when <action or condition>, then <expected outcome>.

## Success indicators

| Indicator or signal | Current situation, if known | Expected outcome | How to track |
|---|---|---|---|

## Dependencies, risks and open questions

| Type | Description | Impact | Next action or owner |
|---|---|---|---|

## Glossary

| Term | Definition in the context of this initiative |
|---|---|

Include business, process, operations and project terms that may be interpreted differently.

## Points for later validation

Include only limitations, concerns or points noticed during discovery that need independent validation.

Label each item as hypothesis, risk or dependency. Do not present them as decisions made.

</prd-template>

## Delivery and next step

After presenting the full PRD, do not save it or start implementation automatically.

Ask for confirmation:

> Does this PRD correctly represent the demand and the decisions made?

If the person confirms, ask:

> What would you like to do next?
>
> 1. Adjust the PRD.
> 2. Save the PRD to a location you choose.
> 3. Start implementation or execution planning.
> 4. Stop for now, without saving.

Before ending, say:

> Since conversations may have a context limit, copy or save this PRD before starting a new session. If you want to continue later without saving it somewhere, use the PRD as the starting point in the new conversation.

Only save or start the next step after an explicit instruction from the person.

If the person chooses to save:

- Ask where and in what format, if that is not defined.
- Preserve the content of the approved PRD.
- Do not change, summarize or reorganize the document without permission.

If the person chooses to implement or plan execution:

- Use the approved PRD as the source of truth.
- Explain that technical decisions, architecture and the development plan must be handled in the next step by the appropriate people or tools.
