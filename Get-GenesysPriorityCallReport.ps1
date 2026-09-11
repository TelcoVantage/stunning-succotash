<#
.SYNOPSIS
    Genesys Cloud - Inbound call priority audit report (CSV).

.DESCRIPTION
    For every inbound voice call that queued in a division, this script produces one CSV row per
    queue attempt showing:

      * the priority Genesys assigned to the call (conversationRoutingData.priority)
      * the requested skills / language and the routing method used
      * how long it waited, whether it was answered/abandoned, who answered and when
      * at the moment the call ENTERED the queue:
            - how many queue members were On Queue / Idle / Interacting / Not Responding
            - how many of the Idle agents were actually ELIGIBLE (hold every requested skill + language)
            - how many other calls were already waiting in that queue (and how many of them had
              higher-or-equal priority, i.e. legitimately ahead of this one)
            - how many other calls were being handled in that queue, and division-wide waiting count
      * while the call was WAITING:
            - how many seconds at least one eligible agent sat Idle, and the longest such stretch
            - how many other calls in the same queue were answered
            - how many of those "jumped ahead" (lower priority, or same priority but entered later),
              and whether the agent who took them was eligible for THIS call
      * a REVIEW flag + reason when something looks like a routing/priority problem.

    Data sources (all read-only):
      POST /api/v2/analytics/conversations/details/query   call list, queue + agent segments
      GET  /api/v2/conversations/{id}                      priority / requested skills / language
      POST /api/v2/analytics/users/details/query           historical agent routing status (IDLE etc.)
      GET  /api/v2/routing/queues/{id}/members             who is a member of each queue
      GET  /api/v2/users?id=...&expand=skills,languages    agent names, skills, languages
      GET  /api/v2/authorization/divisions, /routing/queues/{id}, /routing/skills, /routing/languages

    CONSTRAINED LANGUAGE MODE SAFE (Windows PowerShell 5.1 / AppLocker / WDAC):
      - no .NET static method calls, no ::new(), no [pscustomobject] casts
      - Base64 for the OAuth Basic header is built in pure PowerShell with bit operators
      - only cmdlets, hashtables, arrays, string methods, core-type casts and operators are used

.PARAMETER ClientId / ClientSecret
    OAuth client credentials grant client. Required permissions (role assigned to the client, in the
    division(s) you report on):
        analytics:conversationDetail:view, analytics:userDetail:view, conversation:communication:view,
        routing:queue:view, routing:skill:view, routing:language:view, directory:user:view,
        authorization:division:view
    Optional: by default the credentials embedded at the top of the script ($EmbeddedClientId /
    $EmbeddedClientSecret) are used, then the environment variables GENESYS_CLIENT_ID /
    GENESYS_CLIENT_SECRET. Passing the parameters overrides both.

.PARAMETER Region
    Genesys Cloud region domain, e.g. mypurecloud.com, mypurecloud.ie, mypurecloud.com.au,
    mypurecloud.de, mypurecloud.jp, usw2.pure.cloud, cac1.pure.cloud, euw2.pure.cloud, apne2.pure.cloud,
    aps1.pure.cloud, sae1.pure.cloud, use2.us-gov-pure.cloud, mec1.pure.cloud

.PARAMETER DivisionName / DivisionId
    Division to report on. One of the two is required.

.PARAMETER StartDate / EndDate
    Local date/time range (EndDate exclusive). Default: last 7 days.

.PARAMETER QueueNames
    Optional list of queue names (wildcards allowed) to restrict the report to.

.PARAMETER OutputPath
    CSV output path. Default: .\GenesysPriorityCallReport_<timestamp>.csv

.PARAMETER ChunkHours
    Size of the analytics query windows (default 24h). Keeps every query well inside API limits.

.PARAMETER FlagIdleStretchSeconds
    Flag a row for review when an eligible agent was Idle for at least this many consecutive
    seconds while the call was waiting (default 15). Short idle blips (1-5 s) are normal while the
    platform is selecting/alerting an agent.

.PARAMETER MaxNamesPerCell
    Maximum agent names listed in a single CSV cell (default 15).

.PARAMETER ApiBaseUrl / LoginTokenUrl
    Optional overrides of https://api.<Region> and https://login.<Region>/oauth/token (for proxies or testing).

.EXAMPLE
    # credentials embedded in the script
    .\Get-GenesysPriorityCallReport.ps1 -Region mypurecloud.ie -DivisionName 'Customer Service' `
        -StartDate '2026-09-01' -EndDate '2026-09-08' -OutputPath C:\Temp\PriorityAudit.csv

.EXAMPLE
    .\Get-GenesysPriorityCallReport.ps1 -Region mypurecloud.com -DivisionId 8b1c... -QueueNames 'Sales*','VIP Line'
#>
[CmdletBinding()]
param(
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$Region = 'mypurecloud.com',
    [string]$DivisionName,
    [string]$DivisionId,
    [datetime]$StartDate,
    [datetime]$EndDate,
    [string[]]$QueueNames,
    [string]$OutputPath,
    [int]$ChunkHours = 24,
    [int]$FlagIdleStretchSeconds = 15,
    [int]$MaxNamesPerCell = 15,
    [string]$ApiBaseUrl,
    [string]$LoginTokenUrl
)

$ErrorActionPreference = 'Stop'

# =====================================================================================================
# EMBEDDED CREDENTIALS - fill these in. They are used unless -ClientId / -ClientSecret are passed.
# Anyone who can read this file can call the Genesys API with these credentials: keep the file
# access-controlled and give the OAuth client a read-only role (see README, section 1).
# =====================================================================================================
$EmbeddedClientId     = 'PASTE-YOUR-CLIENT-ID-HERE'
$EmbeddedClientSecret = 'PASTE-YOUR-CLIENT-SECRET-HERE'
$EmbeddedRegion       = ''   # optional, e.g. 'mypurecloud.ie' - overrides the -Region default when set

# =====================================================================================================
# region Helpers (all CLM-safe)
# =====================================================================================================

function ConvertTo-Base64Text {
    # Pure-PowerShell Base64 for ASCII text (used for the OAuth Basic auth header).
    param([string]$Text)
    $alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
    $bytes = @()
    foreach ($ch in $Text.ToCharArray()) {
        $code = [int][char]$ch
        if ($code -gt 127) { throw "ConvertTo-Base64Text only supports ASCII text (found character code $code)." }
        $bytes += $code
    }
    $out = ''
    $i = 0
    $n = $bytes.Count
    while ($i -lt $n) {
        $b0 = $bytes[$i]
        $has1 = (($i + 1) -lt $n)
        $has2 = (($i + 2) -lt $n)
        $b1 = 0
        $b2 = 0
        if ($has1) { $b1 = $bytes[$i + 1] }
        if ($has2) { $b2 = $bytes[$i + 2] }
        $c0 = $b0 -shr 2
        $c1 = (($b0 -band 3) -shl 4) -bor ($b1 -shr 4)
        $c2 = (($b1 -band 15) -shl 2) -bor ($b2 -shr 6)
        $c3 = $b2 -band 63
        $out += $alphabet.Substring($c0, 1)
        $out += $alphabet.Substring($c1, 1)
        if ($has1) { $out += $alphabet.Substring($c2, 1) } else { $out += '=' }
        if ($has2) { $out += $alphabet.Substring($c3, 1) } else { $out += '=' }
        $i += 3
    }
    return $out
}

function ConvertTo-UtcDate {
    # Accepts an ISO-8601 string OR a DateTime (Windows PowerShell 5.1 ConvertFrom-Json already
    # converts ISO strings to local DateTime objects) and returns a UTC DateTime, or $null.
    param($Value)
    if ($Value -eq $null) { return $null }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime() }
    $s = [string]$Value
    if ($s -eq '') { return $null }
    return ([datetime]$s).ToUniversalTime()
}

function Format-IsoUtc {
    param([datetime]$Utc)
    return $Utc.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
}

function Format-LocalTime {
    param($Utc)
    if ($Utc -eq $null) { return '' }
    return $Utc.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
}

function Get-SecondsBetween {
    param($From, $To)
    if ($From -eq $null -or $To -eq $null) { return $null }
    return [int](($To - $From).TotalSeconds)
}

function Get-AbsSeconds {
    param($A, $B)
    $s = ($A - $B).TotalSeconds
    if ($s -lt 0) { $s = 0 - $s }
    return $s
}

function Join-Capped {
    param($Items, [int]$Max)
    $arr = @($Items)
    if ($arr.Count -eq 0) { return '' }
    if ($arr.Count -le $Max) { return ($arr -join '; ') }
    $head = @($arr[0..($Max - 1)])
    return (($head -join '; ') + ('; ... (+' + ($arr.Count - $Max) + ' more)'))
}

function Get-HttpStatusFromError {
    param($ErrorRecord)
    $status = 0
    try { $status = [int]$ErrorRecord.Exception.Response.StatusCode } catch { $status = 0 }
    if ($status -eq 0) {
        $msg = [string]$ErrorRecord.Exception.Message
        if ($msg -like '*429*' -or $msg -like '*Too Many*') { $status = 429 }
        elseif ($msg -like '*404*' -or $msg -like '*Not Found*') { $status = 404 }
        elseif ($msg -like '*401*' -or $msg -like '*Unauthorized*') { $status = 401 }
        elseif ($msg -like '*403*' -or $msg -like '*Forbidden*') { $status = 403 }
        elseif ($msg -like '*400*' -or $msg -like '*Bad Request*') { $status = 400 }
        elseif ($msg -like '*50?*') { $status = 500 }
    }
    return $status
}

function Invoke-GcApi {
    # Wrapper around Invoke-RestMethod with bearer auth, JSON body, 429/5xx retry and optional 404 tolerance.
    param(
        [string]$Method,
        [string]$Path,
        $Body,
        [switch]$AllowNotFound,
        [int]$MaxRetries = 6
    )
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $callParams = @{
                Method      = $Method
                Uri         = ($script:ApiBase + $Path)
                Headers     = @{ Authorization = ('Bearer ' + $script:AccessToken) }
                ContentType = 'application/json'
            }
            if ($Body -ne $null) { $callParams['Body'] = ($Body | ConvertTo-Json -Depth 25 -Compress) }
            $script:ApiCallCount++
            return (Invoke-RestMethod @callParams)
        }
        catch {
            $status = Get-HttpStatusFromError -ErrorRecord $_
            $msg = [string]$_.Exception.Message
            if ($status -eq 404 -and $AllowNotFound) { return $null }
            if (($status -eq 429 -or $status -ge 500) -and $attempt -le $MaxRetries) {
                $wait = 5 * $attempt
                if ($status -eq 429) { $wait = 10 * $attempt }
                Write-Warning ("HTTP {0} on {1} {2} - retrying in {3}s (attempt {4}/{5})" -f $status, $Method, $Path, $wait, $attempt, $MaxRetries)
                Start-Sleep -Seconds $wait
                continue
            }
            if ($status -eq 401) { throw "Unauthorized (401) calling $Method $Path - token expired or invalid. $msg" }
            if ($status -eq 403) { throw "Forbidden (403) calling $Method $Path - the OAuth client is missing a permission (see README). $msg" }
            throw ("Genesys API call failed: {0} {1} HTTP {2}: {3}" -f $Method, $Path, $status, $msg)
        }
    }
}

function Get-GcPagedEntities {
    # Generic pager for GET endpoints returning { entities, pageCount, pageNumber, total }.
    param([string]$PathWithoutPaging, [int]$PageSize = 100, [switch]$AllowNotFound)
    $all = @()
    $page = 1
    while ($true) {
        $sep = '?'
        if ($PathWithoutPaging.Contains('?')) { $sep = '&' }
        $resp = Invoke-GcApi -Method Get -Path ($PathWithoutPaging + $sep + 'pageSize=' + $PageSize + '&pageNumber=' + $page) -AllowNotFound:$AllowNotFound
        if ($resp -eq $null) { return $null }
        $ents = @()
        if ($resp.entities -ne $null) { $ents = @($resp.entities) }
        $all += $ents
        $pageCount = 0
        if ($resp.pageCount -ne $null) { $pageCount = [int]$resp.pageCount }
        if ($ents.Count -lt $PageSize) { break }
        if ($pageCount -gt 0 -and $page -ge $pageCount) { break }
        if ($pageCount -eq 0 -and $ents.Count -eq 0) { break }
        $page++
        if ($page -gt 500) { break }
    }
    return $all
}

function Find-FirstIntervalEndingAfter {
    # $Intervals: array of objects (Start, End[nullable], Status) sorted by Start, non-overlapping.
    # Returns the index of the first interval whose End is $null or > $T, or -1.
    param($Intervals, [datetime]$T)
    $arr = @($Intervals)
    $lo = 0
    $hi = $arr.Count - 1
    $ans = -1
    while ($lo -le $hi) {
        $mid = ($lo + $hi) -shr 1
        $iv = $arr[$mid]
        if ($iv.End -ne $null -and $iv.End -le $T) { $lo = $mid + 1 }
        else { $ans = $mid; $hi = $mid - 1 }
    }
    return $ans
}

function Get-StatusAt {
    # Routing status of a user at instant $T ('IDLE','INTERACTING','OFF_QUEUE','NOT_RESPONDING','COMMUNICATING' or 'UNKNOWN').
    param($Intervals, [datetime]$T)
    if ($Intervals -eq $null) { return 'UNKNOWN' }
    $arr = @($Intervals)
    if ($arr.Count -eq 0) { return 'UNKNOWN' }
    $idx = Find-FirstIntervalEndingAfter -Intervals $arr -T $T
    if ($idx -lt 0) { return 'UNKNOWN' }
    $iv = $arr[$idx]
    if ($iv.Start -le $T) { return $iv.Status }
    return 'UNKNOWN'
}

function Get-IdleIntervalsInWindow {
    # Returns the user's IDLE intervals clipped to [From, To].
    param($Intervals, [datetime]$From, [datetime]$To)
    $result = @()
    if ($Intervals -eq $null) { return $result }
    $arr = @($Intervals)
    if ($arr.Count -eq 0) { return $result }
    $idx = Find-FirstIntervalEndingAfter -Intervals $arr -T $From
    if ($idx -lt 0) { return $result }
    $i = $idx
    while ($i -lt $arr.Count) {
        $iv = $arr[$i]
        if ($iv.Start -ge $To) { break }
        if ($iv.Status -eq 'IDLE') {
            $s = $iv.Start
            $e = $iv.End
            if ($s -lt $From) { $s = $From }
            if ($e -eq $null -or $e -gt $To) { $e = $To }
            if ($e -gt $s) { $result += (New-Object PSObject -Property @{ Start = $s; End = $e }) }
        }
        $i++
    }
    return $result
}

function Get-UnionCoverage {
    # Merges overlapping intervals and returns total covered seconds + longest single merged stretch.
    param($Intervals)
    $arr = @($Intervals | Sort-Object -Property Start)
    $total = 0.0
    $longest = 0.0
    $curS = $null
    $curE = $null
    foreach ($iv in $arr) {
        if ($curS -eq $null) { $curS = $iv.Start; $curE = $iv.End; continue }
        if ($iv.Start -le $curE) {
            if ($iv.End -gt $curE) { $curE = $iv.End }
        }
        else {
            $len = ($curE - $curS).TotalSeconds
            $total += $len
            if ($len -gt $longest) { $longest = $len }
            $curS = $iv.Start
            $curE = $iv.End
        }
    }
    if ($curS -ne $null) {
        $len = ($curE - $curS).TotalSeconds
        $total += $len
        if ($len -gt $longest) { $longest = $len }
    }
    return (New-Object PSObject -Property @{ TotalSeconds = [int]$total; LongestSeconds = [int]$longest })
}

function Test-AgentEligible {
    # An agent is eligible for a call when they hold every requested skill and the requested language.
    param($User, $SkillIds, $LanguageId)
    if ($User -eq $null) { return $false }
    foreach ($sid in @($SkillIds)) {
        if ($sid -eq $null -or $sid -eq '') { continue }
        if (-not ($User.SkillIds -contains $sid)) { return $false }
    }
    if ($LanguageId -ne $null -and $LanguageId -ne '') {
        if (-not ($User.LanguageIds -contains $LanguageId)) { return $false }
    }
    return $true
}

# endregion

# =====================================================================================================
# region Parameters / auth
# =====================================================================================================

if (-not $ClientId -and $EmbeddedClientId -ne '' -and $EmbeddedClientId -notlike 'PASTE-YOUR-*') { $ClientId = $EmbeddedClientId }
if (-not $ClientSecret -and $EmbeddedClientSecret -ne '' -and $EmbeddedClientSecret -notlike 'PASTE-YOUR-*') { $ClientSecret = $EmbeddedClientSecret }
if (-not $ClientId) { $ClientId = $env:GENESYS_CLIENT_ID }
if (-not $ClientSecret) { $ClientSecret = $env:GENESYS_CLIENT_SECRET }
if (-not $ClientId -or -not $ClientSecret) {
    throw 'No credentials: set $EmbeddedClientId / $EmbeddedClientSecret at the top of the script, or pass -ClientId / -ClientSecret.'
}
if ($EmbeddedRegion -ne '' -and -not $PSBoundParameters.ContainsKey('Region')) { $Region = $EmbeddedRegion }
if (-not $DivisionName -and -not $DivisionId) { throw 'Specify -DivisionName or -DivisionId.' }

if ($EndDate -eq $null -or $EndDate -eq [datetime]'0001-01-01') { $EndDate = Get-Date }
if ($StartDate -eq $null -or $StartDate -eq [datetime]'0001-01-01') { $StartDate = $EndDate.AddDays(-7) }
if ($StartDate -ge $EndDate) { throw 'StartDate must be before EndDate.' }
if ($ChunkHours -lt 1) { $ChunkHours = 24 }
if ($ChunkHours -gt 168) { $ChunkHours = 168 }

$StartUtc = $StartDate.ToUniversalTime()
$EndUtc = $EndDate.ToUniversalTime()

if (-not $OutputPath) {
    $OutputPath = Join-Path (Get-Location) ('GenesysPriorityCallReport_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.csv')
}

$script:ApiBase = 'https://api.' + $Region
if ($ApiBaseUrl) { $script:ApiBase = $ApiBaseUrl.TrimEnd('/') }
$tokenUrl = 'https://login.' + $Region + '/oauth/token'
if ($LoginTokenUrl) { $tokenUrl = $LoginTokenUrl }
$script:AccessToken = ''
$script:ApiCallCount = 0

Write-Host ('Genesys Cloud priority call audit  |  region: {0}  |  window (local): {1} -> {2}' -f $Region, $StartDate.ToString('yyyy-MM-dd HH:mm'), $EndDate.ToString('yyyy-MM-dd HH:mm'))

# ---- OAuth client credentials ---------------------------------------------------------------------
Write-Host 'Authenticating...'
$basic = ConvertTo-Base64Text -Text ($ClientId + ':' + $ClientSecret)
$tokenResponse = $null
try {
    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUrl `
        -Headers @{ Authorization = ('Basic ' + $basic) } `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body 'grant_type=client_credentials'
}
catch {
    $st = Get-HttpStatusFromError -ErrorRecord $_
    throw ("OAuth token request failed (HTTP {0}). Check ClientId/ClientSecret/Region. {1}" -f $st, [string]$_.Exception.Message)
}
if ($tokenResponse -eq $null -or -not $tokenResponse.access_token) {
    throw 'OAuth token response did not contain an access_token.'
}
$script:AccessToken = [string]$tokenResponse.access_token
Write-Host 'Authenticated OK.'

# ---- Division -------------------------------------------------------------------------------------
$divisions = @(Get-GcPagedEntities -PathWithoutPaging '/api/v2/authorization/divisions' -PageSize 100)
$division = $null
foreach ($d in $divisions) {
    if ($DivisionId -and ([string]$d.id -eq $DivisionId)) { $division = $d; break }
    if ($DivisionName -and ([string]$d.name -eq $DivisionName)) { $division = $d; break }
}
if ($division -eq $null -and $DivisionName) {
    foreach ($d in $divisions) { if ([string]$d.name -like $DivisionName) { $division = $d; break } }
}
if ($division -eq $null) {
    $names = @($divisions | ForEach-Object { [string]$_.name }) -join ', '
    throw ("Division not found. Available divisions: {0}" -f $names)
}
$DivisionId = [string]$division.id
Write-Host ('Division: {0} ({1})' -f $division.name, $DivisionId)

# ---- Skill / language name maps ------------------------------------------------------------------
$SkillNames = @{}
$LanguageNames = @{}
try {
    foreach ($s in @(Get-GcPagedEntities -PathWithoutPaging '/api/v2/routing/skills' -PageSize 100)) { $SkillNames[[string]$s.id] = [string]$s.name }
    foreach ($l in @(Get-GcPagedEntities -PathWithoutPaging '/api/v2/routing/languages' -PageSize 100)) { $LanguageNames[[string]$l.id] = [string]$l.name }
}
catch {
    Write-Warning ('Could not load skill/language names (ids will be shown instead): ' + [string]$_.Exception.Message)
}

function Get-SkillName { param([string]$Id) if ($SkillNames.ContainsKey($Id)) { return $SkillNames[$Id] } return $Id }
function Get-LanguageName { param([string]$Id) if ($Id -eq '') { return '' } if ($LanguageNames.ContainsKey($Id)) { return $LanguageNames[$Id] } return $Id }

# endregion

# =====================================================================================================
# region 1. Conversation details (inbound voice calls in the division)
# =====================================================================================================

Write-Host 'Querying conversation details...'
$conversationsById = @{}
$chunkStart = $StartUtc
$chunkIndex = 0
while ($chunkStart -lt $EndUtc) {
    $chunkEnd = $chunkStart.AddHours($ChunkHours)
    if ($chunkEnd -gt $EndUtc) { $chunkEnd = $EndUtc }
    $chunkIndex++
    $interval = (Format-IsoUtc $chunkStart) + '/' + (Format-IsoUtc $chunkEnd)
    $page = 1
    $fetched = 0
    while ($true) {
        Write-Progress -Activity 'Conversation details' -Status ("window {0}  page {1}  ({2} calls so far)" -f $interval, $page, $conversationsById.Count)
        $body = @{
            interval            = $interval
            order               = 'asc'
            orderBy             = 'conversationStart'
            conversationFilters = @(
                @{
                    type       = 'and'
                    predicates = @(
                        @{ type = 'dimension'; dimension = 'divisionId'; operator = 'matches'; value = $DivisionId },
                        @{ type = 'dimension'; dimension = 'originatingDirection'; operator = 'matches'; value = 'inbound' }
                    )
                }
            )
            segmentFilters      = @(
                @{
                    type       = 'and'
                    predicates = @(
                        @{ type = 'dimension'; dimension = 'mediaType'; operator = 'matches'; value = 'voice' },
                        @{ type = 'dimension'; dimension = 'direction'; operator = 'matches'; value = 'inbound' }
                    )
                }
            )
            paging              = @{ pageSize = 100; pageNumber = $page }
        }
        $resp = Invoke-GcApi -Method Post -Path '/api/v2/analytics/conversations/details/query' -Body $body
        $convs = @()
        if ($resp -ne $null -and $resp.conversations -ne $null) { $convs = @($resp.conversations) }
        foreach ($c in $convs) { $conversationsById[[string]$c.conversationId] = $c }
        $fetched += $convs.Count
        $totalHits = 0
        if ($resp -ne $null -and $resp.totalHits -ne $null) { $totalHits = [int]$resp.totalHits }
        if ($convs.Count -lt 100) { break }
        if ($totalHits -gt 0 -and $fetched -ge $totalHits) { break }
        $page++
        if ($page -gt 1000) { Write-Warning 'Stopped paging at 1000 pages for one window - reduce -ChunkHours.'; break }
    }
    $chunkStart = $chunkEnd
}
Write-Progress -Activity 'Conversation details' -Completed
Write-Host ('  {0} inbound voice conversations found.' -f $conversationsById.Count)

# ---- Build one "queue attempt" per ACD interact segment --------------------------------------------
$attempts = @()
$queueIds = @{}
$agentIds = @{}
foreach ($convId in $conversationsById.Keys) {
    $conv = $conversationsById[$convId]
    $convStart = ConvertTo-UtcDate $conv.conversationStart
    $convEnd = ConvertTo-UtcDate $conv.conversationEnd

    $ani = ''
    $dnis = ''
    $agentSegments = @()     # (UserId, Type, Start, End)
    $acdSessions = @()

    foreach ($p in @($conv.participants)) {
        $purpose = [string]$p.purpose
        foreach ($sess in @($p.sessions)) {
            if ([string]$sess.mediaType -ne 'voice') { continue }
            if (($purpose -eq 'customer' -or $purpose -eq 'external') -and $ani -eq '') {
                if ($sess.ani -ne $null) { $ani = [string]$sess.ani }
                if ($sess.dnis -ne $null) { $dnis = [string]$sess.dnis }
            }
            if ($purpose -eq 'acd') {
                $acdSessions += (New-Object PSObject -Property @{ Participant = $p; Session = $sess })
            }
            elseif ($purpose -eq 'agent' -or $purpose -eq 'user') {
                $uid = [string]$p.userId
                if ($uid -eq '') { continue }
                foreach ($seg in @($sess.segments)) {
                    $agentSegments += (New-Object PSObject -Property @{
                        UserId = $uid
                        Type   = [string]$seg.segmentType
                        Start  = (ConvertTo-UtcDate $seg.segmentStart)
                        End    = (ConvertTo-UtcDate $seg.segmentEnd)
                    })
                }
            }
        }
    }
    if ($ani -eq '') {
        foreach ($p in @($conv.participants)) { foreach ($sess in @($p.sessions)) { if ($sess.ani -ne $null -and $ani -eq '') { $ani = [string]$sess.ani; $dnis = [string]$sess.dnis } } }
    }

    $attemptNo = 0
    foreach ($acd in $acdSessions) {
        $sess = $acd.Session
        foreach ($seg in @($sess.segments)) {
            if ([string]$seg.segmentType -ne 'interact') { continue }
            $qid = [string]$seg.queueId
            if ($qid -eq '') { $qid = [string]$acd.Participant.queueId }
            if ($qid -eq '') { continue }
            $qs = ConvertTo-UtcDate $seg.segmentStart
            $qe = ConvertTo-UtcDate $seg.segmentEnd
            if ($qs -eq $null) { continue }
            if ($qe -eq $null) { $qe = $EndUtc }
            $attemptNo++
            $queueIds[$qid] = $true

            # Agent answer: earliest agent 'interact' segment starting between queue entry and queue exit (+10s tolerance)
            $answerSeg = $null
            $alertStart = $null
            $offeredNotAnswered = @{}
            $limit = $qe.AddSeconds(10)
            foreach ($as in $agentSegments) {
                if ($as.Start -eq $null) { continue }
                if ($as.Start -lt $qs.AddSeconds(-1) -or $as.Start -gt $limit) { continue }
                if ($as.Type -eq 'interact') {
                    if ($answerSeg -eq $null -or $as.Start -lt $answerSeg.Start) { $answerSeg = $as }
                }
            }
            $answeredBy = ''
            $answerTime = $null
            $handleEnd = $null
            if ($answerSeg -ne $null) {
                $answeredBy = $answerSeg.UserId
                $answerTime = $answerSeg.Start
                $handleEnd = $answerSeg.End
                foreach ($as in $agentSegments) {
                    if ($as.UserId -ne $answeredBy -or $as.Start -eq $null) { continue }
                    if ($as.Type -eq 'interact' -and $as.Start -ge $answerTime -and $as.End -ne $null -and ($handleEnd -eq $null -or $as.End -gt $handleEnd)) { $handleEnd = $as.End }
                    if ($as.Type -eq 'alert' -and $as.Start -ge $qs.AddSeconds(-1) -and $as.Start -le $answerTime) {
                        if ($alertStart -eq $null -or $as.Start -lt $alertStart) { $alertStart = $as.Start }
                    }
                }
            }
            elseif ($sess.selectedAgentId -ne $null -and [string]$sess.selectedAgentId -ne '') {
                # fallback: analytics says an agent was selected but no interact segment was found
                $answeredBy = [string]$sess.selectedAgentId
                $answerTime = $qe
                $handleEnd = $convEnd
            }
            if ($handleEnd -eq $null -and $answerTime -ne $null) { $handleEnd = $convEnd }
            if ($handleEnd -eq $null -and $answerTime -ne $null) { $handleEnd = $answerTime }

            # Agents alerted during the wait who did NOT answer (RONA / declined)
            foreach ($as in $agentSegments) {
                if ($as.Type -ne 'alert' -or $as.Start -eq $null) { continue }
                if ($as.Start -lt $qs.AddSeconds(-1) -or $as.Start -gt $limit) { continue }
                if ($as.UserId -ne $answeredBy) { $offeredNotAnswered[$as.UserId] = $true }
            }
            if ($answeredBy -ne '') { $agentIds[$answeredBy] = $true }
            foreach ($k in $offeredNotAnswered.Keys) { $agentIds[$k] = $true }

            $skillIds = @()
            if ($sess.requestedRoutingSkillIds -ne $null) { foreach ($x in @($sess.requestedRoutingSkillIds)) { $skillIds += [string]$x } }
            $langId = ''
            if ($sess.requestedLanguageId -ne $null) { $langId = [string]$sess.requestedLanguageId }
            $usedRouting = ''
            if ($sess.usedRouting -ne $null) { $usedRouting = [string]$sess.usedRouting }
            $requestedRoutings = ''
            if ($sess.requestedRoutings -ne $null) { $requestedRoutings = (@($sess.requestedRoutings) -join '|') }
            $routingRing = ''
            if ($sess.routingRing -ne $null) { $routingRing = [string]$sess.routingRing }
            $flowOut = ''
            if ($sess.flowOutType -ne $null) { $flowOut = [string]$sess.flowOutType }
            $disc = ''
            if ($seg.disconnectType -ne $null) { $disc = [string]$seg.disconnectType }

            $outcome = 'NotAnswered'
            if ($answeredBy -ne '') { $outcome = 'Answered' }
            elseif ($disc -eq 'client' -or $disc -eq 'peer') { $outcome = 'Abandoned' }
            elseif ($flowOut -ne '') { $outcome = 'FlowOut (' + $flowOut + ')' }
            elseif ($disc -ne '') { $outcome = 'NotAnswered (' + $disc + ')' }

            $attempts += (New-Object PSObject -Property @{
                ConversationId     = $convId
                ConversationStart  = $convStart
                Ani                = $ani
                Dnis               = $dnis
                QueueId            = $qid
                AttemptNo          = $attemptNo
                QueueStart         = $qs
                QueueEnd           = $qe
                WaitSeconds        = (Get-SecondsBetween $qs $qe)
                Outcome            = $outcome
                DisconnectType     = $disc
                FlowOutType        = $flowOut
                AnsweredBy         = $answeredBy
                AnswerTime         = $answerTime
                AlertStart         = $alertStart
                HandleEnd          = $handleEnd
                OfferedNotAnswered = @($offeredNotAnswered.Keys)
                SkillIds           = $skillIds
                LanguageId         = $langId
                UsedRouting        = $usedRouting
                RequestedRoutings  = $requestedRoutings
                RoutingRing        = $routingRing
                Priority           = $null
                PrioritySource     = ''
                PreferredAgents    = ''
            })
        }
    }
}
Write-Host ('  {0} queue attempts across {1} queue(s).' -f $attempts.Count, $queueIds.Count)
if ($attempts.Count -eq 0) {
    Write-Warning 'No queued inbound voice calls found for this division/date range. Nothing to export.'
    return
}

# ---- Queue names + optional queue filter -----------------------------------------------------------
$QueueNameMap = @{}
foreach ($qid in @($queueIds.Keys)) {
    $q = Invoke-GcApi -Method Get -Path ('/api/v2/routing/queues/' + $qid) -AllowNotFound
    if ($q -ne $null -and $q.name -ne $null) { $QueueNameMap[$qid] = [string]$q.name } else { $QueueNameMap[$qid] = $qid + ' (deleted?)' }
}
if ($QueueNames -and $QueueNames.Count -gt 0) {
    $keep = @{}
    foreach ($qid in @($QueueNameMap.Keys)) {
        foreach ($pattern in $QueueNames) { if ($QueueNameMap[$qid] -like $pattern) { $keep[$qid] = $true } }
    }
    $attempts = @($attempts | Where-Object { $keep.ContainsKey($_.QueueId) })
    $queueIds = $keep
    Write-Host ('  Queue filter applied: {0} attempts in {1} queue(s) remain.' -f $attempts.Count, $queueIds.Count)
    if ($attempts.Count -eq 0) { Write-Warning 'No attempts match -QueueNames.'; return }
}

# endregion

# =====================================================================================================
# region 2. Priority per call (GET /api/v2/conversations/{id})
# =====================================================================================================

Write-Host 'Fetching priority / routing data per conversation...'
$attemptsByConv = @{}
foreach ($a in $attempts) {
    if (-not $attemptsByConv.ContainsKey($a.ConversationId)) { $attemptsByConv[$a.ConversationId] = @() }
    $attemptsByConv[$a.ConversationId] += $a
}
$convCounter = 0
$convTotal = $attemptsByConv.Count
$notFound = 0
foreach ($convId in @($attemptsByConv.Keys)) {
    $convCounter++
    if (($convCounter % 25) -eq 0 -or $convCounter -eq $convTotal) {
        Write-Progress -Activity 'Conversation priority lookup' -Status ("{0} / {1}" -f $convCounter, $convTotal) -PercentComplete ([int](100 * $convCounter / $convTotal))
    }
    $detail = Invoke-GcApi -Method Get -Path ('/api/v2/conversations/' + $convId) -AllowNotFound
    if ($detail -eq $null) { $notFound++; foreach ($a in $attemptsByConv[$convId]) { $a.PrioritySource = 'conversation not found (purged?)' }; continue }

    $acdParts = @()
    foreach ($p in @($detail.participants)) {
        if ([string]$p.purpose -eq 'acd') { $acdParts += $p }
    }
    foreach ($a in $attemptsByConv[$convId]) {
        $best = $null
        $bestDiff = 999999999
        foreach ($p in $acdParts) {
            $pq = [string]$p.queueId
            if ($pq -eq '' -and $p.conversationRoutingData -ne $null -and $p.conversationRoutingData.queue -ne $null) { $pq = [string]$p.conversationRoutingData.queue.id }
            if ($pq -ne '' -and $pq -ne $a.QueueId) { continue }
            $pt = $null
            if ($p.connectedTime -ne $null) { $pt = ConvertTo-UtcDate $p.connectedTime }
            if ($pt -eq $null -and $p.startTime -ne $null) { $pt = ConvertTo-UtcDate $p.startTime }
            $diff = 0
            if ($pt -ne $null) { $diff = Get-AbsSeconds $pt $a.QueueStart }
            if ($best -eq $null -or $diff -lt $bestDiff) { $best = $p; $bestDiff = $diff }
        }
        if ($best -eq $null -and $acdParts.Count -gt 0) { $best = $acdParts[0] }
        if ($best -ne $null -and $best.conversationRoutingData -ne $null) {
            $crd = $best.conversationRoutingData
            if ($crd.priority -ne $null) { $a.Priority = [int]$crd.priority; $a.PrioritySource = 'conversationRoutingData' }
            else { $a.Priority = 0; $a.PrioritySource = 'default (no priority on routing data)' }
            if ($a.SkillIds.Count -eq 0 -and $crd.skills -ne $null) {
                $ids = @()
                foreach ($sk in @($crd.skills)) { if ($sk.id -ne $null) { $ids += [string]$sk.id; if ($sk.name -ne $null) { $SkillNames[[string]$sk.id] = [string]$sk.name } } }
                $a.SkillIds = $ids
            }
            if ($a.LanguageId -eq '' -and $crd.language -ne $null -and $crd.language.id -ne $null) {
                $a.LanguageId = [string]$crd.language.id
                if ($crd.language.name -ne $null) { $LanguageNames[$a.LanguageId] = [string]$crd.language.name }
            }
            if ($crd.scoredAgents -ne $null) {
                $names = @()
                foreach ($sa in @($crd.scoredAgents)) {
                    if ($sa.agent -ne $null -and $sa.agent.id -ne $null) { $names += ([string]$sa.agent.id + '=' + [string]$sa.score); $agentIds[[string]$sa.agent.id] = $true }
                }
                $a.PreferredAgents = ($names -join '|')
            }
        }
        else {
            $a.Priority = 0
            $a.PrioritySource = 'no ACD participant routing data'
        }
    }
}
Write-Progress -Activity 'Conversation priority lookup' -Completed
if ($notFound -gt 0) { Write-Warning ("{0} conversation(s) could not be fetched (404) - priority shown as blank." -f $notFound) }

# endregion

# =====================================================================================================
# region 3. Queue members, users (skills/languages) and historical routing status
# =====================================================================================================

Write-Host 'Loading queue members...'
$QueueMembers = @{}      # queueId -> array of userIds
$allUserIds = @{}
foreach ($qid in @($queueIds.Keys)) {
    $members = Get-GcPagedEntities -PathWithoutPaging ('/api/v2/routing/queues/' + $qid + '/members') -PageSize 100 -AllowNotFound
    $ids = @()
    if ($members -ne $null) {
        foreach ($m in @($members)) {
            $uid = ''
            if ($m.id -ne $null) { $uid = [string]$m.id }
            if ($uid -eq '' -and $m.user -ne $null -and $m.user.id -ne $null) { $uid = [string]$m.user.id }
            if ($uid -ne '') { $ids += $uid; $allUserIds[$uid] = $true }
        }
    }
    else {
        Write-Warning ('Queue ' + $QueueNameMap[$qid] + ' not found - member statistics will be blank for it.')
    }
    $QueueMembers[$qid] = $ids
    Write-Host ('  {0}: {1} member(s)' -f $QueueNameMap[$qid], $ids.Count)
}
foreach ($k in @($agentIds.Keys)) { $allUserIds[$k] = $true }

Write-Host ('Loading {0} user(s) with skills/languages...' -f $allUserIds.Count)
$Users = @{}            # userId -> PSObject(Id, Name, SkillIds, LanguageIds)
$idList = @($allUserIds.Keys)
$batchSize = 50
$pos = 0
while ($pos -lt $idList.Count) {
    $endPos = $pos + $batchSize - 1
    if ($endPos -ge $idList.Count) { $endPos = $idList.Count - 1 }
    $batch = @($idList[$pos..$endPos])
    $qs = ''
    foreach ($id in $batch) { $qs += '&id=' + $id }
    $resp = Invoke-GcApi -Method Get -Path ('/api/v2/users?pageSize=' + $batchSize + '&state=any&expand=skills,languages' + $qs)
    if ($resp -ne $null -and $resp.entities -ne $null) {
        foreach ($u in @($resp.entities)) {
            $sk = @()
            $lg = @()
            if ($u.skills -ne $null) { foreach ($s in @($u.skills)) { if ($s.id -ne $null) { $sk += [string]$s.id; if ($s.name -ne $null -and -not $SkillNames.ContainsKey([string]$s.id)) { $SkillNames[[string]$s.id] = [string]$s.name } } } }
            if ($u.languages -ne $null) { foreach ($l in @($u.languages)) { if ($l.id -ne $null) { $lg += [string]$l.id; if ($l.name -ne $null -and -not $LanguageNames.ContainsKey([string]$l.id)) { $LanguageNames[[string]$l.id] = [string]$l.name } } } }
            $Users[[string]$u.id] = (New-Object PSObject -Property @{ Id = [string]$u.id; Name = [string]$u.name; SkillIds = $sk; LanguageIds = $lg })
        }
    }
    $pos = $endPos + 1
}
foreach ($id in $idList) {
    if (-not $Users.ContainsKey($id)) { $Users[$id] = (New-Object PSObject -Property @{ Id = $id; Name = $id; SkillIds = @(); LanguageIds = @() }) }
}
function Get-UserName { param([string]$Id) if ($Id -eq '') { return '' } if ($Users.ContainsKey($Id)) { return $Users[$Id].Name } return $Id }

# ---- Historical routing status for every queue member ------------------------------------------------
$memberIds = @()
foreach ($qid in @($QueueMembers.Keys)) { foreach ($uid in @($QueueMembers[$qid])) { if (-not ($memberIds -contains $uid)) { $memberIds += $uid } } }
Write-Host ('Loading historical routing status for {0} queue member(s)...' -f $memberIds.Count)

$rawIntervals = @{}   # key userId|startIso -> object
$chunkStart = $StartUtc
while ($chunkStart -lt $EndUtc -and $memberIds.Count -gt 0) {
    $chunkEnd = $chunkStart.AddHours($ChunkHours)
    if ($chunkEnd -gt $EndUtc) { $chunkEnd = $EndUtc }
    $interval = (Format-IsoUtc $chunkStart) + '/' + (Format-IsoUtc $chunkEnd)
    $pos = 0
    while ($pos -lt $memberIds.Count) {
        $endPos = $pos + $batchSize - 1
        if ($endPos -ge $memberIds.Count) { $endPos = $memberIds.Count - 1 }
        $batch = @($memberIds[$pos..$endPos])
        $preds = @()
        foreach ($uid in $batch) { $preds += @{ type = 'dimension'; dimension = 'userId'; operator = 'matches'; value = $uid } }
        $page = 1
        while ($true) {
            Write-Progress -Activity 'Agent routing status history' -Status ("window {0}  users {1}-{2}  page {3}" -f $interval, ($pos + 1), ($endPos + 1), $page)
            $body = @{
                interval    = $interval
                userFilters = @(@{ type = 'or'; predicates = $preds })
                paging      = @{ pageSize = 100; pageNumber = $page }
            }
            $resp = Invoke-GcApi -Method Post -Path '/api/v2/analytics/users/details/query' -Body $body
            $details = @()
            if ($resp -ne $null -and $resp.userDetails -ne $null) { $details = @($resp.userDetails) }
            foreach ($ud in $details) {
                $uid = [string]$ud.userId
                foreach ($rs in @($ud.routingStatus)) {
                    $st = ConvertTo-UtcDate $rs.startTime
                    if ($st -eq $null) { continue }
                    $en = ConvertTo-UtcDate $rs.endTime
                    $key = $uid + '|' + (Format-IsoUtc $st)
                    $existing = $null
                    if ($rawIntervals.ContainsKey($key)) { $existing = $rawIntervals[$key] }
                    if ($existing -eq $null -or ($existing.End -ne $null -and ($en -eq $null -or $en -gt $existing.End))) {
                        $rawIntervals[$key] = (New-Object PSObject -Property @{ UserId = $uid; Start = $st; End = $en; Status = [string]$rs.routingStatus })
                    }
                }
            }
            if ($details.Count -lt 100) { break }
            $page++
            if ($page -gt 200) { break }
        }
        $pos = $endPos + 1
    }
    $chunkStart = $chunkEnd
}
Write-Progress -Activity 'Agent routing status history' -Completed

$StatusByUser = @{}   # userId -> sorted array of intervals
$tmp = @{}
foreach ($iv in $rawIntervals.Values) {
    if (-not $tmp.ContainsKey($iv.UserId)) { $tmp[$iv.UserId] = @() }
    $tmp[$iv.UserId] += $iv
}
foreach ($uid in @($tmp.Keys)) { $StatusByUser[$uid] = @($tmp[$uid] | Sort-Object -Property Start) }
Write-Host ('  {0} routing-status interval(s) loaded for {1} user(s).' -f $rawIntervals.Count, $StatusByUser.Count)

# endregion

# =====================================================================================================
# region 4. Correlate: other calls + agent availability at the time of each call
# =====================================================================================================

Write-Host 'Correlating calls with agent availability...'
$sorted = @($attempts | Sort-Object -Property QueueStart)

# Look-back bound: longest wait or handle time in the data set (so the scan windows stay small).
$maxSpanSeconds = 60
foreach ($a in $sorted) {
    $w = Get-SecondsBetween $a.QueueStart $a.QueueEnd
    if ($w -ne $null -and $w -gt $maxSpanSeconds) { $maxSpanSeconds = $w }
    if ($a.AnswerTime -ne $null -and $a.HandleEnd -ne $null) {
        $h = Get-SecondsBetween $a.AnswerTime $a.HandleEnd
        if ($h -ne $null -and $h -gt $maxSpanSeconds) { $maxSpanSeconds = $h }
    }
}

$rows = @()
$rowCounter = 0
for ($idx = 0; $idx -lt $sorted.Count; $idx++) {
    $a = $sorted[$idx]
    $rowCounter++
    if (($rowCounter % 25) -eq 0 -or $rowCounter -eq $sorted.Count) {
        Write-Progress -Activity 'Correlating' -Status ("{0} / {1}" -f $rowCounter, $sorted.Count) -PercentComplete ([int](100 * $rowCounter / $sorted.Count))
    }
    $T = $a.QueueStart
    $Tend = $a.QueueEnd
    if ($a.AnswerTime -ne $null -and $a.AnswerTime -lt $Tend) { $Tend = $a.AnswerTime }
    $lookback = $T.AddSeconds(0 - $maxSpanSeconds - 5)

    # ---- other calls ---------------------------------------------------------------------------
    $waitingSameQueue = 0
    $waitingAheadHigherOrEqual = 0
    $handledSameQueue = 0
    $waitingDivision = 0
    $answeredDuringWait = 0
    $jumped = @()
    $jumpedEligible = 0

    $j = $idx - 1
    while ($j -ge 0) {
        $o = $sorted[$j]
        if ($o.QueueStart -lt $lookback) { break }
        $j--
        if ($o.ConversationId -eq $a.ConversationId -and $o.AttemptNo -eq $a.AttemptNo) { continue }
        $same = ($o.QueueId -eq $a.QueueId)
        # waiting at entry: entered before T and left after T
        if ($o.QueueStart -le $T -and $o.QueueEnd -gt $T) {
            $waitingDivision++
            if ($same) {
                $waitingSameQueue++
                if ($o.Priority -ne $null -and $a.Priority -ne $null -and $o.Priority -ge $a.Priority) { $waitingAheadHigherOrEqual++ }
            }
        }
        if ($same -and $o.AnswerTime -ne $null -and $o.HandleEnd -ne $null -and $o.AnswerTime -le $T -and $o.HandleEnd -gt $T) { $handledSameQueue++ }
        if ($same -and $o.AnswerTime -ne $null -and $o.AnswerTime -gt $T -and $o.AnswerTime -lt $Tend) {
            $answeredDuringWait++
            # entered earlier than us: only a "jump" if it had LOWER priority
            if ($o.Priority -ne $null -and $a.Priority -ne $null -and $o.Priority -lt $a.Priority) {
                $elig = Test-AgentEligible -User $Users[$o.AnsweredBy] -SkillIds $a.SkillIds -LanguageId $a.LanguageId
                if ($elig) { $jumpedEligible++ }
                $jumped += ('{0} (prio {1}, entered {2}, answered {3} by {4}, agentEligibleForThisCall={5})' -f $o.ConversationId, $o.Priority, (Format-LocalTime $o.QueueStart), (Format-LocalTime $o.AnswerTime), (Get-UserName $o.AnsweredBy), $elig)
            }
        }
    }
    $j = $idx + 1
    while ($j -lt $sorted.Count) {
        $o = $sorted[$j]
        if ($o.QueueStart -ge $Tend) { break }
        $j++
        if ($o.ConversationId -eq $a.ConversationId -and $o.AttemptNo -eq $a.AttemptNo) { continue }
        $same = ($o.QueueId -eq $a.QueueId)
        if ($o.QueueStart -le $T -and $o.QueueEnd -gt $T) {
            # identical entry timestamp
            $waitingDivision++
            if ($same) { $waitingSameQueue++; if ($o.Priority -ne $null -and $a.Priority -ne $null -and $o.Priority -ge $a.Priority) { $waitingAheadHigherOrEqual++ } }
        }
        if ($same -and $o.AnswerTime -ne $null -and $o.AnswerTime -gt $T -and $o.AnswerTime -lt $Tend) {
            $answeredDuringWait++
            # entered AFTER us: a jump if lower OR equal priority (FIFO within the same priority)
            if ($o.Priority -ne $null -and $a.Priority -ne $null -and $o.Priority -le $a.Priority) {
                $elig = Test-AgentEligible -User $Users[$o.AnsweredBy] -SkillIds $a.SkillIds -LanguageId $a.LanguageId
                if ($elig) { $jumpedEligible++ }
                $jumped += ('{0} (prio {1}, entered {2}, answered {3} by {4}, agentEligibleForThisCall={5})' -f $o.ConversationId, $o.Priority, (Format-LocalTime $o.QueueStart), (Format-LocalTime $o.AnswerTime), (Get-UserName $o.AnsweredBy), $elig)
            }
        }
    }

    # ---- agents at queue entry / during wait -----------------------------------------------------
    $members = @()
    if ($QueueMembers.ContainsKey($a.QueueId)) { $members = @($QueueMembers[$a.QueueId]) }
    $haveMemberData = ($members.Count -gt 0)
    $onQueue = 0; $idle = 0; $idleEligible = 0; $interacting = 0; $notResponding = 0; $communicating = 0; $unknown = 0
    $idleNames = @()
    $idleEligibleNames = @()
    $idleEligibleAtExit = 0
    $idleAtExit = 0
    $idleWindows = @()
    foreach ($uid in $members) {
        $ivs = $null
        if ($StatusByUser.ContainsKey($uid)) { $ivs = $StatusByUser[$uid] }
        $st = Get-StatusAt -Intervals $ivs -T $T
        $u = $Users[$uid]
        $elig = Test-AgentEligible -User $u -SkillIds $a.SkillIds -LanguageId $a.LanguageId
        switch ($st) {
            'IDLE' {
                $onQueue++; $idle++
                $idleNames += $u.Name
                if ($elig) { $idleEligible++; $idleEligibleNames += $u.Name }
            }
            'INTERACTING' { $onQueue++; $interacting++ }
            'COMMUNICATING' { $onQueue++; $communicating++ }
            'NOT_RESPONDING' { $onQueue++; $notResponding++ }
            'OFF_QUEUE' { }
            default { $unknown++ }
        }
        $stExit = Get-StatusAt -Intervals $ivs -T $Tend
        if ($stExit -eq 'IDLE') { $idleAtExit++; if ($elig) { $idleEligibleAtExit++ } }
        if ($elig -and $Tend -gt $T) { $idleWindows += @(Get-IdleIntervalsInWindow -Intervals $ivs -From $T -To $Tend) }
    }
    $coverage = Get-UnionCoverage -Intervals $idleWindows

    # ---- review flag -----------------------------------------------------------------------------
    $reasons = @()
    if ($jumpedEligible -gt 0) { $reasons += ('{0} lower/later-priority call(s) answered by an eligible agent while this call waited' -f $jumpedEligible) }
    elseif ($jumped.Count -gt 0) { $reasons += ('{0} lower/later-priority call(s) answered first (by agents NOT eligible for this call - check skills/language)' -f $jumped.Count) }
    if ($haveMemberData -and $coverage.LongestSeconds -ge $FlagIdleStretchSeconds) { $reasons += ('an eligible agent was Idle for {0}s continuously while this call waited' -f $coverage.LongestSeconds) }
    if ($a.OfferedNotAnswered.Count -gt 0) { $reasons += ('offered to {0} agent(s) who did not answer' -f $a.OfferedNotAnswered.Count) }
    if ($haveMemberData -and $a.Outcome -eq 'Abandoned' -and $idleEligible -gt 0) { $reasons += 'abandoned although eligible agents were Idle at queue entry' }
    $flag = ''
    if ($reasons.Count -gt 0) { $flag = 'REVIEW' }

    $skillNamesText = @()
    foreach ($sid in $a.SkillIds) { $skillNamesText += (Get-SkillName $sid) }
    $ronaNames = @()
    foreach ($uid in $a.OfferedNotAnswered) { $ronaNames += (Get-UserName $uid) }

    $prioText = ''
    if ($a.Priority -ne $null) { $prioText = $a.Priority }
    $na = ''
    if (-not $haveMemberData) { $na = 'n/a' }

    $rows += (New-Object PSObject -Property @{
        ReviewFlag                              = $flag
        ReviewReason                            = ($reasons -join ' | ')
        ConversationId                          = $a.ConversationId
        ConversationStartLocal                  = (Format-LocalTime $a.ConversationStart)
        ANI                                     = $a.Ani
        DNIS                                    = $a.Dnis
        QueueName                               = $QueueNameMap[$a.QueueId]
        QueueId                                 = $a.QueueId
        QueueAttempt                            = $a.AttemptNo
        Priority                                = $prioText
        PrioritySource                          = $a.PrioritySource
        RequestedSkills                         = ($skillNamesText -join '; ')
        RequestedLanguage                       = (Get-LanguageName $a.LanguageId)
        RoutingMethodUsed                       = $a.UsedRouting
        RoutingMethodsRequested                 = $a.RequestedRoutings
        BullseyeRing                            = $a.RoutingRing
        PreferredAgents                         = $a.PreferredAgents
        QueueEntryTimeLocal                     = (Format-LocalTime $a.QueueStart)
        QueueExitTimeLocal                      = (Format-LocalTime $a.QueueEnd)
        WaitSeconds                             = $a.WaitSeconds
        Outcome                                 = $a.Outcome
        AnsweredBy                              = (Get-UserName $a.AnsweredBy)
        AnswerTimeLocal                         = (Format-LocalTime $a.AnswerTime)
        FirstAlertTimeLocal                     = (Format-LocalTime $a.AlertStart)
        AlertToAnswerSeconds                    = (Get-SecondsBetween $a.AlertStart $a.AnswerTime)
        OfferedButNotAnsweredBy                 = (Join-Capped $ronaNames $MaxNamesPerCell)
        QueueMembersTotal                       = $members.Count
        AgentsOnQueueAtEntry                    = $(if ($haveMemberData) { $onQueue } else { $na })
        AgentsIdleAtEntry                       = $(if ($haveMemberData) { $idle } else { $na })
        AgentsIdleAndEligibleAtEntry            = $(if ($haveMemberData) { $idleEligible } else { $na })
        IdleEligibleAgentNamesAtEntry           = (Join-Capped $idleEligibleNames $MaxNamesPerCell)
        IdleAgentNamesAtEntry                   = (Join-Capped $idleNames $MaxNamesPerCell)
        AgentsInteractingAtEntry                = $(if ($haveMemberData) { $interacting } else { $na })
        AgentsCommunicatingAtEntry              = $(if ($haveMemberData) { $communicating } else { $na })
        AgentsNotRespondingAtEntry              = $(if ($haveMemberData) { $notResponding } else { $na })
        AgentsStatusUnknownAtEntry              = $(if ($haveMemberData) { $unknown } else { $na })
        AgentsIdleAtExit                        = $(if ($haveMemberData) { $idleAtExit } else { $na })
        AgentsIdleAndEligibleAtExit             = $(if ($haveMemberData) { $idleEligibleAtExit } else { $na })
        SecondsAnEligibleAgentWasIdleDuringWait = $(if ($haveMemberData) { $coverage.TotalSeconds } else { $na })
        LongestEligibleIdleStretchSeconds       = $(if ($haveMemberData) { $coverage.LongestSeconds } else { $na })
        OtherCallsWaitingInQueueAtEntry         = $waitingSameQueue
        OtherCallsAheadWithHigherOrEqualPriority = $waitingAheadHigherOrEqual
        OtherCallsBeingHandledInQueueAtEntry    = $handledSameQueue
        DivisionCallsWaitingAtEntry             = $waitingDivision
        CallsAnsweredInQueueDuringWait          = $answeredDuringWait
        CallsJumpedAhead                        = $jumped.Count
        CallsJumpedAheadByEligibleAgent         = $jumpedEligible
        JumpedAheadDetail                       = ($jumped -join ' || ')
        DisconnectType                          = $a.DisconnectType
    })
}
Write-Progress -Activity 'Correlating' -Completed

# endregion

# =====================================================================================================
# region 5. Export + summary
# =====================================================================================================

$columns = @(
    'ReviewFlag', 'ReviewReason',
    'ConversationId', 'ConversationStartLocal', 'ANI', 'DNIS',
    'QueueName', 'QueueAttempt', 'Priority', 'PrioritySource',
    'RequestedSkills', 'RequestedLanguage', 'RoutingMethodUsed', 'RoutingMethodsRequested', 'BullseyeRing', 'PreferredAgents',
    'QueueEntryTimeLocal', 'QueueExitTimeLocal', 'WaitSeconds', 'Outcome',
    'AnsweredBy', 'AnswerTimeLocal', 'FirstAlertTimeLocal', 'AlertToAnswerSeconds', 'OfferedButNotAnsweredBy',
    'QueueMembersTotal', 'AgentsOnQueueAtEntry', 'AgentsIdleAtEntry', 'AgentsIdleAndEligibleAtEntry',
    'IdleEligibleAgentNamesAtEntry', 'IdleAgentNamesAtEntry',
    'AgentsInteractingAtEntry', 'AgentsCommunicatingAtEntry', 'AgentsNotRespondingAtEntry', 'AgentsStatusUnknownAtEntry',
    'AgentsIdleAtExit', 'AgentsIdleAndEligibleAtExit',
    'SecondsAnEligibleAgentWasIdleDuringWait', 'LongestEligibleIdleStretchSeconds',
    'OtherCallsWaitingInQueueAtEntry', 'OtherCallsAheadWithHigherOrEqualPriority', 'OtherCallsBeingHandledInQueueAtEntry',
    'DivisionCallsWaitingAtEntry', 'CallsAnsweredInQueueDuringWait',
    'CallsJumpedAhead', 'CallsJumpedAheadByEligibleAgent', 'JumpedAheadDetail',
    'DisconnectType', 'QueueId'
)

$rows | Sort-Object -Property QueueEntryTimeLocal | Select-Object $columns | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

$flagged = @($rows | Where-Object { $_.ReviewFlag -eq 'REVIEW' })
$answered = @($rows | Where-Object { $_.Outcome -eq 'Answered' })
$abandoned = @($rows | Where-Object { $_.Outcome -eq 'Abandoned' })
Write-Host ''
Write-Host '================ SUMMARY ================'
Write-Host ('Division            : {0}' -f $division.name)
Write-Host ('Window (local)      : {0} -> {1}' -f $StartDate.ToString('yyyy-MM-dd HH:mm'), $EndDate.ToString('yyyy-MM-dd HH:mm'))
Write-Host ('Queue attempts      : {0}  (answered {1}, abandoned {2})' -f $rows.Count, $answered.Count, $abandoned.Count)
Write-Host ('Flagged for review  : {0}' -f $flagged.Count)
$byPrio = @($rows | Group-Object -Property Priority | Sort-Object -Property Name)
foreach ($g in $byPrio) {
    $grpFlag = @($g.Group | Where-Object { $_.ReviewFlag -eq 'REVIEW' }).Count
    $avgWait = 0
    if ($g.Count -gt 0) { $avgWait = [int](($g.Group | Measure-Object -Property WaitSeconds -Average).Average) }
    Write-Host ('  priority {0,-4}: {1,5} call(s), avg wait {2,5}s, {3} flagged' -f $g.Name, $g.Count, $avgWait, $grpFlag)
}
Write-Host ('API calls made      : {0}' -f $script:ApiCallCount)
Write-Host ('CSV written         : {0}' -f $OutputPath)
Write-Host '========================================='
Write-Host 'Tip: filter ReviewFlag = REVIEW and read ReviewReason / JumpedAheadDetail first. See README.md for how to interpret the columns.'

# endregion
