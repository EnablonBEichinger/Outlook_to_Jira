param(
    [string]$JsonPath,
    [switch]$PreviewOnly
)

# Every run gets its own transcript log; old calendar exports are pruned in the finally block below.
$logsDir = Join-Path $PSScriptRoot "Jira_Logs"
if (-not (Test-Path -LiteralPath $logsDir)) { New-Item -ItemType Directory -Path $logsDir | Out-Null }
$logPath = Join-Path $logsDir ("Jira_Upload_{0}.txt" -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'))
Start-Transcript -Path $logPath | Out-Null

try {

Write-Host "Run date/time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan

# === Jira Configuration ===
$localConfigPath = Join-Path $PSScriptRoot "config.json"
$legacyConfigPath = "C:\TEMP\config.json"
$configPath = if (Test-Path -LiteralPath $legacyConfigPath) { $legacyConfigPath } else { $localConfigPath }

if (Test-Path -LiteralPath $configPath) {
    $config = Get-Content -Raw -LiteralPath $configPath | ConvertFrom-Json
    $debugMode = $config.debugMode
    $JiraCloudDomain = $config.JiraCloudDomain
    $EmailAddress = $config.EmailAddress
    $AccountID = $config.AccountID
    $API_Token = $config.API_Token
    $ProjectKey = $config.ProjectKey
    $ParentTaskKey = $config.ParentTaskKey
    $Division = $config.Division
    $BusinessUnit = $config.BusinessUnit
    $Components = $config.Components
    $DateOffsetNegative = $config.DateOffsetNegative
}
else {
    Write-Host "Configuration file not found at '$localConfigPath' or '$legacyConfigPath'." -ForegroundColor Red
    exit 1
}

# Auto-detect today's calendar export (named "Calendar_yyyy-MM-dd.json") unless -JsonPath was explicitly supplied.
if (-not $PSBoundParameters.ContainsKey('JsonPath')) {
    $todayFileName = "Calendar_{0}.json" -f (Get-Date -Format 'yyyy-MM-dd')
    $todayPath = Join-Path $PSScriptRoot $todayFileName
    if (Test-Path -LiteralPath $todayPath) {
        $JsonPath = $todayPath
    }
    else {
        $latest = Get-ChildItem -LiteralPath $PSScriptRoot -Filter 'Calendar_*.json' -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
        if ($latest) {
            Write-Host "No export found for today ($todayFileName). Falling back to the most recently modified export." -ForegroundColor Yellow
            $JsonPath = $latest.FullName
        }
        else {
            $JsonPath = $todayPath
        }
    }
}

if (-not (Test-Path -LiteralPath $JsonPath)) {
    Write-Host "Calendar JSON not found: $JsonPath" -ForegroundColor Red
    exit 1
}

Write-Host "Using the file named '$(Split-Path -Leaf $JsonPath)'." -ForegroundColor Cyan

# Counts
$countEventsParsed = 0
$countEventsIgnored = 0
$countJiraTasksCreated = 0
$countJiraTasksIgnored = 0

# Added/ignored entries for the end-of-run log listings.
$addedLog = @()
$ignoredLog = @()

# Debug mode settings
$debugModeLimit = 1  # Set to $null for no limit.

# Define start and end of the current week (Monday to Sunday)
$today = Get-Date
$startOfWeek = $today.AddDays(-($today.DayOfWeek.value__ - 1 + $DateOffsetNegative)).Date
$endOfWeek = $startOfWeek.AddDays(7).AddSeconds(-1)

# Prompt user to confirm start and end of the week
$startOfWeekStr = $startOfWeek.ToString('yyyy-MM-dd HH:mm')
$endOfWeekStr = $endOfWeek.ToString('yyyy-MM-dd HH:mm')
Write-Host "Reading calendar JSON: $JsonPath" -ForegroundColor Cyan
Write-Host "Start of the week: $startOfWeekStr" -ForegroundColor Green
Write-Host "End of the week:   $endOfWeekStr" -ForegroundColor Yellow

# Get ISO 8601 week number
$culture = [System.Globalization.CultureInfo]::InvariantCulture
$calendarInfo = $culture.Calendar
$weekRule = [System.Globalization.CalendarWeekRule]::FirstFourDayWeek
$firstDayOfWeek = [System.DayOfWeek]::Monday
$weekNumber = $calendarInfo.GetWeekOfYear($startOfWeek, $weekRule, $firstDayOfWeek)

# The Outlook export is a single-line JSON array of appointment objects.
$jsonCulture = [System.Globalization.CultureInfo]::InvariantCulture
$rows = Get-Content -Raw -LiteralPath $JsonPath | ConvertFrom-Json
$appointments = @()

foreach ($row in $rows) {
    try {
        $start = [datetime]::Parse($row.start, $jsonCulture)
        $end = [datetime]::Parse($row.end, $jsonCulture)
    }
    catch {
        if (-not $PreviewOnly) { Write-Host "Ignoring event with invalid date/time: $($row.subject)" -ForegroundColor Yellow }
        $ignoredLog += [PSCustomObject]@{ Subject = $row.subject; Start = $null; Reason = 'Invalid date/time' }
        $countEventsIgnored++
        continue
    }

    $categoriesJoined = ($row.categories -join ';')
    $isPrivate = $row.sensitivity -match '^(?i:private|confidential)$'
    # Don't text-match "Private" in subject/categories - it also matches topics like "Private Cloud".
    $ignoreReasons = @()
    if ($start -lt $startOfWeek -or $end -gt $endOfWeek) { $ignoreReasons += 'Outside selected week' }
    if ($isPrivate) { $ignoreReasons += 'Sensitivity=Private' }
    if ($categoriesJoined -like '*Jira Ignore*') { $ignoreReasons += 'Category=Jira Ignore' }
    if ($row.subject -like '*Canceled*') { $ignoreReasons += 'Canceled' }
    if ($row.subject -like '*Declined*' -or $row.responseType -eq 'declined') { $ignoreReasons += 'Declined' }
    if ($row.subject -like '*[[]Team Absence]*') { $ignoreReasons += 'Team Absence' }
    if ($row.isAllDay) { $ignoreReasons += 'All-day event' }

    if ($ignoreReasons.Count -gt 0) {
        if (-not $PreviewOnly) { Write-Host "Ignoring event: $($row.subject)" -ForegroundColor Yellow }
        $ignoredLog += [PSCustomObject]@{ Subject = $row.subject; Start = $start; Reason = ($ignoreReasons -join ', ') }
        $countEventsIgnored++
        continue
    }

    $duration = $end - $start
    $durationFormatted = "{0}h {1}m" -f [math]::Floor($duration.TotalHours), $duration.Minutes
    $totalMinutes = [math]::Round($duration.TotalMinutes)

    # Keep the original story-point mapping.
    if ($totalMinutes -le 15) { $storypoints = .25 }
    elseif ($totalMinutes -le 30) { $storypoints = .5 }
    elseif ($totalMinutes -le 45) { $storypoints = .75 }
    elseif ($totalMinutes -le 60) { $storypoints = 1 }
    elseif ($totalMinutes -le 90) { $storypoints = 1.5 }
    elseif ($totalMinutes -le 120) { $storypoints = 2 }
    elseif ($totalMinutes -le 150) { $storypoints = 2.5 }
    elseif ($totalMinutes -le 180) { $storypoints = 3 }
    elseif ($totalMinutes -le 210) { $storypoints = 3.5 }
    elseif ($totalMinutes -le 240) { $storypoints = 4 }
    else { $storypoints = 0 }

    $dayPrefix = if ($row.subject -match 'Daily') { "$($start.DayOfWeek) - " } else { '' }
    $entryId = "$($row.subject)|$($start.ToString('o'))|$($end.ToString('o'))|$($row.location)"

    $appointments += [PSCustomObject]@{
        EntryID = $entryId
        Categories = $categoriesJoined
        Sensitivity = $row.sensitivity
        Subject = "Week $weekNumber - $dayPrefix$($row.subject)"
        Start = $start
        End = $end
        Duration = $durationFormatted
        Location = $row.location
        StoryPts = $storypoints
    }
    $countEventsParsed++
}

Write-Host "Parsed $countEventsParsed calendar event(s) from JSON." -ForegroundColor Green

function Write-IgnoredLog {
    Write-Host "`n=== Ignored ===" -ForegroundColor Yellow
    if ($ignoredLog.Count -eq 0) {
        Write-Host "None"
    }
    else {
        foreach ($item in $ignoredLog) {
            $startText = if ($item.Start) { $item.Start.ToString('yyyy-MM-dd HH:mm') } else { 'n/a' }
            Write-Host " - $($item.Subject) | $startText | $($item.Reason)"
        }
    }
}

if ($PreviewOnly) {
    Write-Host "`n=== Preview: Events That Would Be Sent to Jira ===" -ForegroundColor Cyan
    if ($appointments.Count -eq 0) {
        Write-Host "No calendar events passed the filters." -ForegroundColor Yellow
    }
    else {
        for ($index = 0; $index -lt $appointments.Count; $index++) {
            $appointment = $appointments[$index]
            Write-Host ("{0,2}. {1} | {2} - {3} | {4} | {5} | {6} story points" -f `
                ($index + 1),
                $appointment.Subject,
                $appointment.Start.ToString('yyyy-MM-dd HH:mm'),
                $appointment.End.ToString('HH:mm'),
                $appointment.Duration,
                $appointment.Categories,
                $appointment.StoryPts)
        }
    }
    Write-Host "Total events that would be sent to Jira: $($appointments.Count)" -ForegroundColor Green
    Write-IgnoredLog
    exit
}

if ($debugMode) { $appointments }

# Encode credentials for Basic Auth
$authString = $EmailAddress + ":" + $API_Token
$Base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($authString))
$decodedAccountID = [System.Net.WebUtility]::UrlDecode([string]$AccountID)
$headers = @{
    Authorization = "Basic $Base64AuthInfo"
    'Content-Type' = 'application/json'
}

foreach ($appt in $appointments) {
    if ($debugMode) {
        Write-Host "Debug Mode is ON." -ForegroundColor Yellow
        if ($debugModeLimit -and ($appointments.IndexOf($appt)) -ge $debugModeLimit) {
            Write-Host "Debug Mode limit of $debugModeLimit reached. Stopping further processing." -ForegroundColor Red
            break
        }
    }

    Write-Host "`nProcessing event: $($appt.Subject)" -ForegroundColor Cyan
    $inputString = "$($appt.EntryID)$($appt.Start.ToString('yyyy-MM-dd HH:mm:ss'))"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($inputString)
    $md5 = [System.Security.Cryptography.MD5]::Create()
    $md5Hash = [BitConverter]::ToString($md5.ComputeHash($bytes)) -replace '-', ''
    $issueTypeName = if ([string]::IsNullOrEmpty($ParentTaskKey)) { 'Task' } else { 'Sub-task' }

    $fields = @{
        project = @{ key = $ProjectKey }
        summary = $appt.Subject
        description = @{ type = 'doc'; version = 1; content = @(@{ type = 'paragraph'; content = @(@{ type = 'text'; text = "MD5 Hash: $md5Hash" }) }) }
        issuetype = @{ name = $issueTypeName }
        assignee = @{ id = $decodedAccountID }
        timetracking = @{ originalEstimate = $appt.Duration; remainingEstimate = 0 }
        priority = @{ name = 'Medium' }
        customfield_10146 = $appt.StoryPts
        customfield_10257 = $appt.Start.ToString('yyyy-MM-dd')
        customfield_10255 = @{ value = $Division }
        customfield_10160 = @{ value = $BusinessUnit }
        duedate = $appt.End.ToString('yyyy-MM-dd')
        labels = @('Meeting', 'Outlook_to_Jira')
    }
    if (-not [string]::IsNullOrEmpty($ParentTaskKey)) { $fields.parent = @{ key = $ParentTaskKey } }
    $jsonBody = (@{ fields = $fields } | ConvertTo-Json -Depth 50)

    Write-Host "`n--- Jira Task Preview ---"
    Write-Host $jsonBody
    Write-Host "--------------------------"
    Write-Host 'Checking for any existing tasks with the same MD5 Hash...'

    try {
        $searchUri = "https://$JiraCloudDomain/rest/api/3/search/jql"
        $body = @{ jql = "project = $ProjectKey AND description ~ 'MD5 Hash: $md5Hash'"; maxResults = 1; fields = @('id') } | ConvertTo-Json
        $response = Invoke-RestMethod -Uri $searchUri -Method Post -Headers $headers -Body $body
        if ($response.issues.Count -gt 0) {
            Write-Host "A task with MD5 Hash $md5Hash already exists in Jira. Skipping creation." -ForegroundColor Yellow
            $ignoredLog += [PSCustomObject]@{ Subject = $appt.Subject; Start = $appt.Start; Reason = 'Duplicate - Jira issue already exists' }
            $countJiraTasksIgnored++
            continue
        }
        Write-Host "No existing task found with MD5 Hash $md5Hash. Proceeding to create a new task." -ForegroundColor Green
    }
    catch {
        Write-Host "Error occurred while checking existing tasks: $_" -ForegroundColor Red
        break
    }

    try {
        $issueUri = "https://$JiraCloudDomain/rest/api/3/issue"
        $response = Invoke-RestMethod -Uri $issueUri -Method Post -Headers $headers -Body $jsonBody
        if (-not $response.id) { Write-Host 'Failed to create task. No Issue ID returned.'; break }
        $issueKey = $response.key
        Write-Host "Task created successfully with Issue Key: $issueKey"
        $addedLog += [PSCustomObject]@{ Subject = $appt.Subject; Start = $appt.Start; IssueKey = $issueKey }
        $countJiraTasksCreated++

        $jiraWorkLog = @{
            comment = @{ type = 'doc'; version = 1; content = @(@{ type = 'paragraph'; content = @(@{ type = 'text'; text = $appt.Subject }) }) }
            started = $appt.Start.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff") + '+0000'
            timeSpent = $appt.Duration
        }
        $worklogBody = $jiraWorkLog | ConvertTo-Json -Depth 50
        $worklogUri = "https://$JiraCloudDomain/rest/api/3/issue/$issueKey/worklog"
        $worklogResponse = Invoke-RestMethod -Uri $worklogUri -Method Post -Headers $headers -Body $worklogBody
        if (-not $worklogResponse.id) { Write-Host 'Failed to add work log. No Worklog ID returned.'; break }
        Write-Host "Work log added successfully with Worklog ID: $($worklogResponse.id)"

        $transitionBody = @{ transition = @{ id = '31' } } | ConvertTo-Json
        $transitionUri = "https://$JiraCloudDomain/rest/api/3/issue/$issueKey/transitions"
        Invoke-RestMethod -Uri $transitionUri -Method Post -Headers $headers -Body $transitionBody | Out-Null
        Write-Host "Issue $issueKey transitioned to Done."
    }
    catch {
        Write-Host "Error occurred while sending request to Jira API: $_" -ForegroundColor Red
        break
    }
}

Write-Host "`n=== Added to Jira ===" -ForegroundColor Green
if ($addedLog.Count -eq 0) {
    Write-Host "None"
}
else {
    foreach ($item in $addedLog) {
        Write-Host " - $($item.Subject) | $($item.Start.ToString('yyyy-MM-dd HH:mm')) | $($item.IssueKey)"
    }
}
Write-IgnoredLog

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "Total JSON Events Parsed: $countEventsParsed" -ForegroundColor Green
Write-Host "Total JSON Events Ignored: $countEventsIgnored" -ForegroundColor Yellow
Write-Host "Total Jira Tasks Created: $countJiraTasksCreated" -ForegroundColor Green
Write-Host "Total Jira Tasks Ignored: $countJiraTasksIgnored" -ForegroundColor Yellow

}
finally {
    # Keep the last 14 days of calendar exports; delete anything older based on the date in the filename.
    $cutoffDate = (Get-Date).Date.AddDays(-14)
    Get-ChildItem -LiteralPath $PSScriptRoot -Filter 'Calendar_*.json' -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.BaseName -match '^Calendar_(\d{4}-\d{2}-\d{2})$') {
            $fileDate = [datetime]::ParseExact($matches[1], 'yyyy-MM-dd', $null)
            if ($fileDate -lt $cutoffDate) {
                Write-Host "Deleting old calendar export (older than 14 days): $($_.Name)" -ForegroundColor DarkGray
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            }
        }
    }
    Stop-Transcript | Out-Null
}