# Example: discovery for the digital delivery check-in

Fictitious session showing how the skill runs the interview. Company, people and numbers are made up. This is the same project used in the charter, risk matrix and status report examples in this repository.

## Initial request

> The queues at the gate are huge. I want to digitize truck check-in.

## Round 1

The initial frontier has four decisions that do not depend on each other.

❓ **Q1. Main problem**: is the pain the waiting time at the gate, the lack of visibility into deliveries or the rework of typing the forms? All three may exist, but each leads to a different outcome.

➡️ I recommend treating waiting time as the main problem, because it is visible to whoever arrives, measurable and what prompted the request.

❓ **Q2. Coverage**: does digital check-in apply to all sites or only one?

➡️ I recommend starting with one site. This allows validating the process before replicating it.

❓ **Q3. Scheduling**: is advance delivery scheduling part of this demand?

➡️ I recommend leaving it out. Scheduling changes the process of carriers and sales, which forms another workstream with its own outcome.

❓ **Q4. Expected outcome**: how long does gate service take today, and what would be a good result?

➡️ I did not find this data in the available materials. If it has not been measured, I recommend measuring for one week before setting the target.

**Answers:**
1. Waiting time.
2. Only the main site.
3. Out of scope.
4. The gate recorded an average of 20 minutes last month. The target is 8.

## Round 2

The answers opened new decisions.

❓ **Q5. Exceptions**: what happens when a truck arrives without a valid invoice?

➡️ I recommend not completing the check-in and sending the driver to the gate service desk, which already handles this case today.

❓ **Q6. Queue visibility**: does the gate team need to see the queue in real time, or is the record enough?

➡️ I recommend including the queue view, because it lets the gate react before waiting time grows.

❓ **Q7. Ready to use**: when will the gate team be able to run the new process?

➡️ I recommend requiring training with an attendance list as an acceptance condition.

**Answers:** 5, 6 and 7 accepted as recommended.

**Scope leak identified:** a question came up about how the warehouse should reorganize receiving based on the queue. That decides how another area reacts to the outcome, so it was recorded under "Out of scope", with the warehouse as owner.

## Consolidated stories

1. As a driver, I need to check in without filling in a paper form, so that I am released faster.
2. As the gate team, I need to see the vehicle queue in real time, so that I can prioritize service.
3. As the gate team, I need check-in to be blocked when the invoice is not valid, so that irregular deliveries are not released.
4. To measure the project outcome, the arrival and release time of each vehicle must be recorded, so that the average service time can be tracked.

**Confirmation:** the person confirmed the four stories, and the PRD was generated.

## Resulting PRD (summary)

# Digital delivery check-in

## Problem statement

At the main site, truck check-in is done on a paper form and typed in later by the gate team. Service takes 20 minutes on average, according to the gate's records for the last month, which creates queues and delays receiving.

## Objective and expected outcome

Reduce the average gate service time from 20 to 8 minutes, sustained for 30 consecutive days.

## Scope

### Included

- Digital check-in at the main site gate.
- Real-time queue view for the gate.
- Gate team training.

### Out of scope

- **Delivery scheduling:** changes the process of carriers and sales and should be handled as its own initiative.
- **Rollout to other sites:** will be assessed after validation at this site.
- **Reorganization of warehouse receiving:** a warehouse decision, based on the queue data.

## Rules and acceptance criteria

| ID | Condition | Expected outcome | Exceptions or notes |
|---|---|---|---|
| BR01 | Given the vehicle has no valid invoice, when the driver tries to complete check-in | Check-in is not completed and the driver is sent to the service desk | None |
| AC01 | Given the 30-day period after go-live | Average service time is up to 8 minutes | Measured by arrival and release times |

## Points for later validation

- **Hypothesis:** Wi-Fi coverage at the gate is enough for digital check-in.
