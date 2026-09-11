# Genesys Cloud – Inbound Call Priority Audit Report

`Get-GenesysPriorityCallReport.ps1` exports a CSV with **one row per queue attempt** for every inbound
voice call that queued in a division, and answers the question a contact-centre manager usually asks:

> *"At the moment this priority call came in, was anybody idle, and did another call get answered first?"*

For every call it records the priority Genesys assigned, the requested skills/language, who was idle
(and who was idle **and eligible**) at queue entry and during the wait, how many other calls were
waiting/being handled, and which lower-priority calls were answered before it ("jumped ahead").
Rows that look wrong are flagged `REVIEW` with a plain-English reason.

The script is **Constrained Language Mode safe** (Windows PowerShell 5.1, AppLocker/WDAC): no .NET
static calls, no `::new()`, no `[pscustomobject]`; the OAuth Basic header is Base64-encoded in pure
PowerShell with bit operators.

---

## 1. Prerequisites

### OAuth client (Client Credentials grant)
Admin > Integrations > OAuth > *Add client* > grant type **Client Credentials**. Assign a role that
holds these permissions **in the division you report on** (or all divisions):

| Permission | Used for |
|---|---|
| `analytics:conversationDetail:view` | conversation details (call list, queue + agent segments) |
| `analytics:userDetail:view` | historical agent routing status (Idle / Interacting / Not Responding) |
| `conversation:communication:view` | `GET /api/v2/conversations/{id}` – priority, skills, language |
| `routing:queue:view` | queue names and queue membership |
| `routing:skill:view`, `routing:language:view` | skill / language names |
| `directory:user:view` | agent names, skills, languages |
| `authorization:division:view` | division lookup |

### Embedded settings
Open `Get-GenesysPriorityCallReport.ps1` and fill in the block at the top:

```powershell
$EmbeddedClientId     = 'PASTE-YOUR-CLIENT-ID-HERE'
$EmbeddedClientSecret = 'PASTE-YOUR-CLIENT-SECRET-HERE'
$EmbeddedRegion       = 'mypurecloud.com.au'          # Australia (Sydney)
$EmbeddedDivisionId   = 'PASTE-YOUR-DIVISION-ID-HERE'
$DefaultReportDays    = 5                             # default window: last 5 days
```

Anyone who can read the file can use those credentials, so restrict access to the file and keep the
OAuth client's role read-only. Command-line parameters (`-ClientId`, `-ClientSecret`, `-Region`,
`-DivisionId`/`-DivisionName`, `-StartDate`/`-EndDate`) override the embedded values when given.

The division id is the GUID shown in the URL on Admin > Account Settings > Divisions, or from
`GET /api/v2/authorization/divisions` in the API Explorer.

Other regions, if ever needed: `mypurecloud.com`, `mypurecloud.ie`, `mypurecloud.de`, `mypurecloud.jp`,
`usw2.pure.cloud`, `cac1.pure.cloud`, `euw2.pure.cloud`, `apne2.pure.cloud`, `aps1.pure.cloud`,
`sae1.pure.cloud`, `mec1.pure.cloud`.

---

## 2. Running it

With everything embedded, the last 5 days for the embedded division:

```powershell
.\Get-GenesysPriorityCallReport.ps1
```

Or with an explicit window and output file:

```powershell
.\Get-GenesysPriorityCallReport.ps1 -StartDate '2026-09-01' -EndDate '2026-09-08' -OutputPath 'C:\Temp\PriorityAudit.csv'
```

Useful options

| Parameter | Default | Notes |
|---|---|---|
| `-DivisionId` / `-DivisionName` | embedded division id | report on a different division |
| `-Region` | `mypurecloud.com.au` | Genesys region domain |
| `-QueueNames 'VIP*','Sales'` | all queues | wildcard filter on queue name |
| `-StartDate` / `-EndDate` | last 5 days | local time, `EndDate` exclusive |
| `-FlagIdleStretchSeconds` | 15 | flag when an eligible agent was idle this long *continuously* during the wait |
| `-ChunkHours` | 24 | analytics query window size (keeps every query inside API limits) |
| `-MaxNamesPerCell` | 15 | cap for agent-name lists in a cell |
| `-EventsOutputPath` | `<OutputPath>_PriorityEvents.csv` | where the priority event log CSV is written |
| `-ClientId` / `-ClientSecret` | embedded values | override the credentials embedded at the top of the script |

Runtime: one `GET /api/v2/conversations/{id}` per call is needed for the priority, so a busy division
over a month may take a while (progress bars are shown; 429 rate limits are retried automatically).
All times in the CSV are in the **local time zone of the machine running the script**.

---

## 3. Reading the CSV

Start by filtering `ReviewFlag = REVIEW` and reading `ReviewReason`, `PriorityEvidence` and
`JumpedAheadDetail`. For the "did priority work" question, open the second CSV (`…_PriorityEvents.csv`,
section below) and filter on `Verdict`.

### Identity / priority
| Column | Meaning |
|---|---|
| `ConversationId`, `ConversationStartLocal`, `ANI`, `DNIS` | the call |
| `QueueName`, `QueueAttempt` | queue and attempt number (a call transferred/re-queued has several rows) |
| `Priority` | `conversationRoutingData.priority` of the ACD participant – the priority the queue used for this call (final value, i.e. after any *Set Priority* / escalation in the flow). Blank when the conversation is no longer retrievable (see `PrioritySource`). |
| `RequestedSkills`, `RequestedLanguage` | ACD skills / language the flow requested – an agent needs **all** of them to be offered the call |
| `RoutingMethodUsed`, `RoutingMethodsRequested`, `BullseyeRing`, `PreferredAgents` | how the call was routed (Standard, Bullseye ring n, Preferred agent, Predictive, Manual) |

### Timing / outcome
| Column | Meaning |
|---|---|
| `QueueEntryTimeLocal`, `QueueExitTimeLocal`, `WaitSeconds` | time in queue |
| `Outcome` | `Answered`, `Abandoned`, `FlowOut (…)` (queue timeout → voicemail/callback/other flow) or `NotAnswered (disconnectType)` |
| `AnsweredBy`, `AnswerTimeLocal`, `FirstAlertTimeLocal`, `AlertToAnswerSeconds` | who answered and how long it rang |
| `OfferedButNotAnsweredBy` | agents the call was alerted to who did not pick up (RONA / declined) |

### Agents at the moment the call entered the queue (`…AtEntry`)
Counts are over the queue's **current** members (Genesys does not keep historical membership).
Status comes from the analytics routing-status history, so it is exactly what Genesys saw at that instant.

| Column | Meaning |
|---|---|
| `QueueMembersTotal` | members of the queue |
| `AgentsOnQueueAtEntry` | members whose routing status was not Off Queue (Idle + Interacting + Communicating + Not Responding) |
| `AgentsIdleAtEntry`, `IdleAgentNamesAtEntry` | members in **Idle** (on queue, no interaction) |
| `AgentsIdleAndEligibleAtEntry`, `IdleEligibleAgentNamesAtEntry` | of those, the ones that also hold every requested skill and the language – **only these could have been offered the call** |
| `AgentsInteractingAtEntry`, `AgentsCommunicatingAtEntry`, `AgentsNotRespondingAtEntry` | the rest of the on-queue members |
| `AgentsStatusUnknownAtEntry` | members with no routing-status record at that instant (never on queue that day, or user data not returned) |
| `AgentsIdleAtExit`, `AgentsIdleAndEligibleAtExit` | same two idle counts at the moment the call left the queue (answered/abandoned) |

### Agents while the call waited
| Column | Meaning |
|---|---|
| `SecondsAnEligibleAgentWasIdleDuringWait` | number of seconds during the wait when **at least one eligible member was Idle** |
| `LongestEligibleIdleStretchSeconds` | longest continuous such stretch. 1–5 s blips are normal (that is the platform picking and alerting an agent); tens of seconds while the call waits is not |

### Other calls
| Column | Meaning |
|---|---|
| `OtherCallsWaitingInQueueAtEntry` | calls already waiting in the same queue when this one entered |
| `OtherCallsAheadWithHigherOrEqualPriority` | of those, the ones legitimately ahead (higher priority, or same priority and older) |
| `OtherCallsBeingHandledInQueueAtEntry` | calls from this queue being handled by an agent at that moment |
| `DivisionCallsWaitingAtEntry` | calls waiting in *any* queue of the division at that moment |
| `CallsAnsweredInQueueDuringWait` | calls from the same queue answered while this call waited |
| `AnsweredBeforeThisCallConversationIds` | the conversation ids of those calls, `;` separated (paste one into the Genesys interaction search to open it) |

### Priority evidence (did priority work?)
Every comparison is between calls **in the same queue**, using the priority Genesys recorded on each call.

| Column | Meaning |
|---|---|
| `PriorityEvidence` | one-line verdict for this row: `HONOURED: …` and/or `VIOLATED: …`, blank when nothing relevant happened while it waited |
| `OvertookLowerPriorityCalls`, `OvertookLowerPriorityConversationIds`, `OvertookLowerPriorityDetail` | **proof priority worked, seen from the high-priority call**: lower-priority calls that entered the queue *before* this one and were still waiting when this one was answered. Example: this row is a 400 answered at 10:20:20; the detail lists the 100 that entered at 10:20:00 and was only answered at 10:20:40. |
| `HigherPriorityCallsServedFirst`, `HigherPriorityServedFirstConversationIds`, `HigherPriorityServedFirstDetail` | **the same proof seen from the low-priority call**: higher-priority calls that entered *after* this one and were answered while it was still waiting. Correct behaviour, not a fault. |
| `CallsJumpedAhead`, `JumpedAheadConversationIds` | calls answered during this call's wait that **should have been behind it**: lower priority, or same priority but entered the queue later |
| `CallsJumpedAheadByEligibleAgent` | jumped-ahead calls answered by an agent who held this call's skills/language – i.e. that agent could have taken *this* call instead. **This is the number that proves or disproves the manager's claim.** |
| `JumpedAheadDetail` | `conversationId (prio, entered, answered by, agentEligibleForThisCall=True/False)` for each jumped-ahead call |

### The priority event log (`…_PriorityEvents.csv`)
A second CSV is written next to the main one (or at `-EventsOutputPath`). It has **one row per pair**
"call that was answered" / "call that was still waiting in the same queue at that moment" whenever the
pair says something about priority. Sorted by time, it is the audit trail to hand to the manager:

| `Verdict` | Meaning |
|---|---|
| `PRIORITY HONOURED` | the answered call had **higher** priority than a call that had entered the queue **earlier** and was still waiting. E.g. `AnsweredConversationId` = a 400 entered 10:20:10, answered 10:20:20; `WaitingConversationId` = a 100 entered 10:20:00, answered 10:20:40. |
| `PRIORITY VIOLATED` | the answered call had **lower** priority than a call that was still waiting. Check `AnsweringAgentEligibleForWaiting`: `False` means the agent lacked the waiting call's skills/language (not a priority fault); `True` with `AnsweredRoutingMethod = Standard` is a genuine violation. |
| `FIFO VIOLATED (same priority)` | equal priority, but a younger call was answered before an older one (again check eligibility). |

Other columns: `AnsweredPriority` / `WaitingPriority`, both calls' queue entry times, `AnsweredAtLocal`,
`AnsweredBy`, how long each had waited at that instant, what eventually happened to the waiting call
(`WaitingOutcome`, `WaitingAnsweredAtLocal`, `WaitingAnsweredBy`), `WaitingRequestedSkills`, and a `Note`
explaining the most likely non-priority cause when one is visible.

The console summary prints the totals of each verdict, so a healthy queue reads e.g.
`PRIORITY HONOURED: 57, PRIORITY VIOLATED: 0`.

### `ReviewFlag` / `ReviewReason`
A row is flagged when any of these is true:
* a lower/later-priority call was answered by an **eligible** agent while this call waited (strong evidence of a priority problem),
* a lower/later-priority call was answered first but by a **non-eligible** agent (usually skills/language, not priority),
* an eligible agent was Idle for at least `-FlagIdleStretchSeconds` continuously while this call waited,
* the call was offered to agents who did not answer,
* the call abandoned while eligible agents were Idle at queue entry.

---

## 4. How Genesys priority actually works (things that look like bugs but are not)

1. **Priority only orders the queue.** A higher priority never pulls an agent off an interaction and never
   interrupts after-call work. If everybody eligible is `Interacting`, the priority call waits.
2. **Skills and language trump priority.** An Idle agent without the requested skills/language is invisible
   to that call. This is the most common reason for "agents were idle but the VIP waited" –
   compare `AgentsIdleAtEntry` with `AgentsIdleAndEligibleAtEntry`.
3. **Bullseye rings** deliberately withhold calls from outer-ring agents for the ring timeout, even when
   they are idle. Check `RoutingMethodUsed` / `BullseyeRing`.
4. **Preferred-agent routing** waits for the scored agents first (`PreferredAgents`).
5. **Agents in several queues** are offered the highest-priority interaction across *all* their queues, then
   the longest waiting. A call in another queue with higher priority can win an agent that is also a
   member of this queue – see `DivisionCallsWaitingAtEntry`.
6. **The offer is not instantaneous.** After an agent turns Idle, Genesys takes a few seconds to select and
   alert; then the call rings (`AlertToAnswerSeconds`). Idle stretches of a few seconds are expected.
7. **RONA**: if the agent does not answer, they go `Not Responding` and the call is re-offered –
   `OfferedButNotAnsweredBy` shows who. The call keeps its position and priority.
8. **The priority shown is the final one.** If the flow raises priority after the call already queued, the
   earlier part of the wait was at the lower priority.
9. **Membership is current, not historical.** An agent removed from the queue after the date range still
   appears as a member (and vice versa) – their status history still reflects what they actually did.

A genuine priority routing fault looks like this in the CSV: `CallsJumpedAheadByEligibleAgent > 0` on a
high-priority row, with `RoutingMethodUsed = Standard`, no Bullseye/Preferred routing, and
`RequestedSkills` satisfied by the agent listed in `JumpedAheadDetail`. If you find such rows, raise a
Genesys Cloud support case quoting the `ConversationId` values.

---

## 5. Test harness (development only)

`tests/mock-genesys-server.js` is a tiny Node.js mock of the endpoints used, with a scenario whose
expected numbers are known. `tests/Run-MockTest.ps1` runs the report against it under
Constrained Language Mode (`$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'`).

```powershell
node tests/mock-genesys-server.js   # in one terminal
pwsh -File tests/Run-MockTest.ps1   # in another
```
