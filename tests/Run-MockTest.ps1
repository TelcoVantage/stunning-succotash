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
#           LongestEligibleIdleStretchSeconds = 28 ; c4 flagged as abandoned while Carol Agent was idle.
Import-Csv $out | Select-Object ConversationId, Priority, ReviewFlag, WaitSeconds, AgentsIdleAtEntry, AgentsIdleAndEligibleAtEntry, CallsJumpedAheadByEligibleAgent, LongestEligibleIdleStretchSeconds | Format-Table -AutoSize
