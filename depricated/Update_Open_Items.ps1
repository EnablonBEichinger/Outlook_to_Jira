# === Jira Config ===
$JiraCloudDomain = "yourdomain.atlassian.net"
$EmailAddress    = "your_email@example.com"
$ApiToken        = "your_api_token"
$ProjectKey      = "PROJ"

# === Build Auth Header ===
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$EmailAddress:$ApiToken"))
$Headers = @{ Authorization = "Basic $base64AuthInfo"; "Content-Type" = "application/json" }

# === Step 1: Get all open issues ===
$JQL = "assignee = currentUser() AND status != Done"
$SearchUri = "https://$JiraCloudDomain/rest/api/2/search"
$Body = @{ jql = $JQL; fields = @("key","timetracking") } | ConvertTo-Json -Depth 10
$Issues = Invoke-RestMethod -Uri $SearchUri -Method Post -Headers $Headers -Body $Body

foreach ($issue in $Issues.issues) {
    $IssueKey = $issue.key
    $OrigEst  = $issue.fields.timetracking.originalEstimateSeconds

    if ($OrigEst -gt 0) {
        # === Step 2: Log work ===
        $WorklogUri = "https://$JiraCloudDomain/rest/api/2/issue/$IssueKey/worklog"
        $WorklogBody = @{ timeSpentSeconds = $OrigEst; comment = "Auto-logged work equal to Original Estimate" } | ConvertTo-Json
        Invoke-RestMethod -Uri $WorklogUri -Method Post -Headers $Headers -Body $WorklogBody
    }

    # === Step 3: Transition to Done ===
    $TransitionUri = "https://$JiraCloudDomain/rest/api/2/issue/$IssueKey/transitions"
    $Transitions = Invoke-RestMethod -Uri $TransitionUri -Method Get -Headers $Headers

    $DoneTransition = $Transitions.transitions | Where-Object { $_.name -eq "Done" }
    if ($DoneTransition) {
        $TransitionBody = @{ transition = @{ id = $DoneTransition.id } } | ConvertTo-Json
        Invoke-RestMethod -Uri $TransitionUri -Method Post -Headers $Headers -Body $TransitionBody
    }
}
