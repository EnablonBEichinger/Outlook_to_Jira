# === Jira Configuration ===
$JiraCloudDomain = ""
$EmailAddress    = ""
$AccountID       = ""
$API_Token       = ""
$ProjectKey      = ""

$API_Specific = "/rest/api/2/issue"
$JiraURI = "https://" + $JiraCloudDomain + $API_Specific

# Create Outlook COM object
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$calendar = $namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar

# Define start and end of the current week (Monday to Sunday)
$today = Get-Date
$startOfWeek = $today.AddDays(-($today.DayOfWeek.value__ - 1)).Date
$endOfWeek = $startOfWeek.AddDays(7).AddSeconds(-1)

# Get ISO 8601 week number
$culture = [System.Globalization.CultureInfo]::InvariantCulture
$calendarInfo = $culture.Calendar
$weekRule = [System.Globalization.CalendarWeekRule]::FirstFourDayWeek
$firstDayOfWeek = [System.DayOfWeek]::Monday
$weekNumber = $calendarInfo.GetWeekOfYear($today, $weekRule, $firstDayOfWeek)

# Get calendar items and include recurring events
$items = $calendar.Items
$items.IncludeRecurrences = $true
$items.Sort("[Start]")

# Define filter for current week
$filter = "[Start] >= '" + $startOfWeek.ToString("g") + "' AND [End] <= '" + $endOfWeek.ToString("g") + "'"
$filteredItems = $items.Restrict($filter)

# Initialize array to hold appointment data
$appointments = @()

# Loop through and collect non-private, non-"Jira Ignore" appointments with duration
foreach ($item in $filteredItems) {
    if ($item.Sensitivity -ne 2 -and ($item.Categories -notlike "*Jira Ignore*")) {
        $start = [datetime]$item.Start
        $end = [datetime]$item.End
        $duration = $end - $start
        $durationFormatted = "{0}h {1}m" -f $duration.Hours, $duration.Minutes
        $totalMinutes = [math]::Round($duration.TotalMinutes)

        # Add day prefix if subject contains "Daily"
        $dayPrefix = ""
        if ($item.Subject -match "Daily") {
            $dayPrefix = "$($start.DayOfWeek) - "
        }

        # Calculate the Story Points
        if($totalMinutes -le 15){$storypoints = .25}
            elseif($totalMinutes -le 30){$storypoints = .5}
            elseif($totalMinutes -le 45){$storypoints = .75}
            elseif($totalMinutes -le 60){$storypoints = 1}
            elseif($totalMinutes -le 90){$storypoints = 1.5}
            elseif($totalMinutes -le 120){$storypoints = 2}
            elseif($totalMinutes -le 150){$storypoints = 2.5}
            elseif($totalMinutes -le 180){$storypoints = 3}
            elseif($totalMinutes -le 210){$storypoints = 3.5}
            elseif($totalMinutes -le 240){$storypoints = 4}
            else{$storypoints = 0}

        # Create a custom object for each appointment
        $appt = [PSCustomObject]@{
            EntryID   = $item.EntryID
            Subject   = "Week $weekNumber - $dayPrefix$($item.Subject)"
            Start     = $start
            End       = $end
            Duration  = $durationFormatted
            Location  = $item.Location
            StoryPts  = $storypoints
        }

        # Add to array
        $appointments += $appt
    }
}

# Output the array (optional)
# $appointments

# Encode credentials for Basic Auth
$authString = $EmailAddress + ":" + $API_Token
$Base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($authString))

# === Loop through Outlook appointments and stage Jira tasks ===
foreach ($appt in $appointments) {
    # Construct Jira issue payload
    $jiraTask = @{
        fields = @{
            project         = @{ key = $ProjectKey }
            summary         = $appt.Subject
            issuetype       = @{ name = "Task" }
            assignee        = @{ id = $AccountID}
            timetracking    = @{ originalEstimate = $appt.Duration }
            priority        = @{ name = "Medium" } 
            customfield_10146 = $appt.StoryPts
        }
    }

    # Convert to JSON for review or sending
    $jsonBody = $jiraTask | ConvertTo-Json -Depth 50

    # Output staged task (for review)
    Write-Host "`n--- Jira Task Preview ---"
    Write-Host $jsonBody
    Write-Host "--------------------------"

    # To actually send to Jira, uncomment the following lines:
    $headers = @{
        "Authorization" = "Basic $Base64AuthInfo"
        "Content-Type"  = "application/json"
    }
    Invoke-RestMethod -Uri "https://wkenterprise.atlassian.net/rest/api/2/issue" -Method Post -Headers $headers -Body $jsonBody
}
