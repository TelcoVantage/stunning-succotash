# Runs the report against tests/mock-genesys-server.js (node tests/mock-genesys-server.js) in Constrained Language Mode.
$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$script = Join-Path (Split-Path -Parent $here) 'Get-GenesysPriorityCallReport.ps1'
$out = Join-Path $here 'mock-output.csv'
& $script -ClientId 'client-id' -ClientSecret 'client-secret' -Region 'mock.local' `
    -ApiBaseUrl 'http://localhost:8091' -LoginTokenUrl 'http://localhost:8091/oauth/token' `
    -DivisionName 'Customer Service' -StartDate '2026-09-02 09:00' -EndDate '2026-09-02 12:00' `
    -OutputPath $out
# Expected: c2 (priority 5) flagged with CallsJumpedAheadByEligibleAgent = 1 (c3 answered by Bob Agent),
#           LongestEligibleIdleStretchSeconds = 28 ; c4 flagged as abandoned while Carol Agent was idle;
#           c7 (priority 400) shows OvertookLowerPriorityCalls = 1 (c6, priority 100) and c6 shows
#           HigherPriorityCallsServedFirst = 1. Events CSV: 1 PRIORITY HONOURED (c7 over c6), 1 PRIORITY VIOLATED (c3 over c2).
#           Agent decisions CSV: 1 CORRECT (Alice took c7/400 while c6/100 waited), 2 WRONG ORDER (Bob took c3/0 while
#           c2/5 waited; Alice took c9/400 in Customer Service Line while c8/800 waited in AU RAS).
#           Queue config CSV: AU RAS flagged (ConversationScore while the shared queue is TimestampAndPriority, bullseye).
Import-Csv $out | Select-Object ConversationId, Priority, ReviewFlag, WaitSeconds, AgentsIdleAndEligibleAtEntry, AnsweredBeforeThisCallConversationIds, OvertookLowerPriorityCalls, HigherPriorityCallsServedFirst, CallsJumpedAheadByEligibleAgent, PriorityEvidence | Format-Table -AutoSize
Import-Csv (Join-Path $here 'mock-output_PriorityEvents.csv') | Select-Object Verdict, AnsweredConversationId, AnsweredPriority, AnsweredAtLocal, WaitingConversationId, WaitingPriority, WaitingEnteredQueueLocal, AnsweringAgentEligibleForWaiting | Format-Table -AutoSize
Import-Csv (Join-Path $here 'mock-output_AgentDecisions.csv') | Select-Object DecisionTimeLocal, Agent, Verdict, TakenConversationId, TakenPriority, ShouldHaveTaken | Format-Table -AutoSize
Import-Csv (Join-Path $here 'mock-output_QueueConfig.csv') | Select-Object QueueName, ScoringMethod, SkillEvaluationMethod, RoutingMethod, Members, SharedAgentsWithQueues, Warning | Format-Table -AutoSize
