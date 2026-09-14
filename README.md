## Overview
The PowerShell scripts here are designed to utilize a JSON export of the users Outlook Calendar.  The process defined will parse that JSON file applying logic and standard processes and add entries to Jira.  In doing so the script will:
- Apply a "Week ##" identifier
- Determine the story points
- Create a HASH to prevent duplicates
- Close the Jira task
- Save a log locally of what was done

## Prerequisites
There are a few items needed to be in place before these scripts can be run, these items are detailed below.
### PowerShell 7
- [https://learn.microsoft.com/en-us/powershell/scripting/windows-powershell/ise/introducing-the-windows-powershell-ise?view=powershell-7.5](https://learn.microsoft.com/en-us/powershell/scripting/windows-powershell/ise/introducing-the-windows-powershell-ise?view=powershell-7.5)
### Power Automate
This is the process that will create the actual JSON file.  As the final step in the Flow it will save the file to a designated location in OneDrive.  The instructions for creating the flow can be found in  Confluence:
https://enablon.atlassian.net/wiki/spaces/ICIint/pages/137875357700/Outlook+to+Jira+-+Power+Automate+Flow
### One Drive 
This needs to be setup on the machine and syncing.  It is preferable to create a folder just for this process but ultimately any location will do.  This location will hold the JSON files (the script cleans the location but holds 14 days worth just in case) as well as a directory for the log files.  The PowerShell script is made to run from this location for ease of use but could be edited to be outside this folder. 

## Config File
To use this script you will first need to copy the `config.json` file, once created you can then populate the config file as follows:

- `debugMode`
    - This is for testing purposes, only use if you know what you're doing.
- `JiraCloudDomain`
    - The main URL used for Jira. For WK, this should be prepopulated in the template.
- `EmailAddress`
    - This is your WK email address (the email used for the Jira account)
- `AccountID`
    - Login to Jira and click to edit your profile
    - In the URL, after "/people/" and up to "?" this is the profile ID
- `API_Token`
    - Instructions for creating API token can be found: [Jira Instructions](https://support.atlassian.com/atlassian-account/docs/manage-api-tokens-for-your-atlassian-account/)
    - Create API token
        - [https://id.atlassian.com/manage-profile/security/api-tokens](https://id.atlassian.com/manage-profile/security/api-tokens)
            - Named "Outlook_to_Jira"
- `ProjectKey`
    - For WK, this should be prepopulated in the template. If not, follow:
        - Click on "View All Projects"
        - Filter for the Project you are working on
        - There is a field called "Key", this is the value you need
- `ParentTaskKey` - Optional
    - This determines if the task created is under a Jira "story", if so popluate with the story key, for example: "CPESG-8365"
- `Division`
    - A "custom" field in Jira. For WK, this should be prepopulated in the template.
- `BusinessUnit`
    - A "custom" field in Jira. For WK, this should be prepopulated in the template.
- `Components`
    - Defines which board the task will appear on, options are:
        - `Enablon Public cloud AWS`
        - `Enablon Public cloud Azure`
        - `Enablon Private cloud`
- `StartDate` - Optional
    - Set to a specific date in `yyyy-MM-dd` format (e.g. `2026-08-03`) to filter from that date instead of the current week.
    - Leave blank (`""`) to use the current week (Monday-Sunday), in which case `EndDate` is ignored.
- `EndDate` - Optional
    - Set to a specific date in `yyyy-MM-dd` format to define the end of the filtered range (inclusive). Requires `StartDate` to also be set.
    - Leave blank (`""`) with `StartDate` set to default to a 7-day window starting from `StartDate`.
- `Silent` - Optional, defaults to `false`
    - When `true`, skips the interactive confirmation prompts (Parent Task Key and date range) so the script can run unattended, e.g. from a scheduled task.

> **Breaking change:** `DateOffsetNegative` has been removed. If your `config.json` still has it, delete it and use `StartDate` instead (e.g. an old offset of 7 becomes a `StartDate` of last Monday's date).

`JiraCloudDomain`, `EmailAddress`, `AccountID`, `API_Token`, `ProjectKey`, `Division`, `BusinessUnit`, and `Components` are required. The script checks for these on startup and exits with a list of any missing settings rather than failing partway through a run.
### Filters
This script will _**NOT**_ process events in Outlook if any of the following are true:
#### Event Sensitivity
- Private
#### Subject contains
- Private
- Declined
#### Outlook Categories
- Jira Ignore
- Private
