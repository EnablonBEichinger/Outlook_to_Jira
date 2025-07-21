# === Jira Configuration ===
$JiraCloudDomain = ""
$EmailAddress    = ""
$AccountID       = ""
$API_Token       = ""
$ProjectKey      = ""
$Assignee         = ""

# Encode credentials for Basic Auth
#$Base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$EmailAddress':'$API_Token"))
$authString = $EmailAddress + ":" + $API_Token
$Base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($authString))

# === Test Data for Jira Staging ===
$appointments = @(
    [PSCustomObject]@{
        Subject   = "Week 29 - Monday - Test01"
        Start     = (Get-Date).Date.AddHours(9)
        End       = (Get-Date).Date.AddHours(9).AddMinutes(30)
        Duration  = "0h 30m"
    },
    [PSCustomObject]@{
        Subject   = "Week 29 - Tuesday - Test02"
        Start     = (Get-Date).Date.AddDays(1).AddHours(9)
        End       = (Get-Date).Date.AddDays(1).AddHours(9).AddMinutes(30)
        Duration  = "0h 30m"
    }
)

# === Loop through Outlook appointments and stage Jira tasks ===
foreach ($appt in $appointments) {
    # Optional: parse duration into minutes if needed
    $durationParts = $appt.Duration -split 'h |m'
    $hours = [int]$durationParts[0]
    $minutes = [int]$durationParts[1]
    $totalMinutes = ($hours * 60) + $minutes

    # Construct Jira issue payload
    $jiraTask = @{
        fields = @{
            project         = @{ key = $ProjectKey }
            summary         = $appt.Subject
            issuetype       = @{ name = "Task" }
            assignee        = @{ id = $AccountID}
            # Optional: add assignee, priority, labels, etc.
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
