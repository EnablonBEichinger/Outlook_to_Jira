## Overview
The PowerShell script here is designed to work against Microsoft Outlook on Windows to extract the current weeks calendar events for population in to Jira.  The script has been designed to use certain flags in Outlook as filters so as to be able to manage what items go to Jira directly from Outlook. 

### Repo Content
Outlook_GatherEventsOnly.ps1 - If you would like to simply get a count of events for the time period.

### Filters
The use of the Private flag in Outlook is a filter in this script.  If a user does not want to transfer a calendar entry to Jira the calendar entry can be marked as Private. 

A "Jira Ignore" category can be added by the user to their Outlook as well to achieve the same result.  Any meeting added to this category will be ignored for import in to Jira.  

### Time Restriction 
While it can be adjusted there is a time range of 1 week for export.  The theory with that is that this script can be set as a scheduled task for the start of the week so that the users entries can automatically be added for the week.  Since its common for meetings to be added throughout the week, this time frame can also logically be changed to even a single day and set with the same Scheduled Task automation so that the user will freshly have their information daily to reduce edits.  

### Be Aware
If setting this script to run as a Scheduled Task, timing is important.  Since it would logically be run from the users machine it should be set for a time when the user will be logged on and Outlook open.  It is not a requirement for Outlook to be open for this script to work, however being open will ensure that the calendar is up to date.   