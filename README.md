## Overview
The PowerShell script here is designed to work against Microsoft Outlook on Windows to extract the current weeks calendar events for population in to Jira.  The script has been designed to use certain flags in Outlook as filters so as to be able to manage what items go to Jira directly from Outlook. 

### Repo Content
- **Outlook_GatherEventsOnly.ps1** - If you would like to simply get a count of events for the time period.
- **Jira_Upload_Testing.ps1** - This is a testing script.  Only a shell of the main section this is here in order to be able to test the variables for Jira access to ensure that items are working properly before the first large push of data. 
- **Outlook_To_Jira.ps1** - This is the weekly sync script.  Running this script will pull all non-filtered calendar entries in to Jira.
- **Outlook_To_Jira_Updates.ps1** - A seperate script adding the ability to do mid-week syncs for new items.  For an item to be added to Jira it needs to be added to the "Jira Update" category.  Be sure to remove it from the category after syncing or it will create a duplicate on the next run. 

### Filters
The use of the Private flag in Outlook is a filter in this script.  If a user does not want to transfer a calendar entry to Jira the calendar entry can be marked as ***Private***. 

Outlook Categories needed: 
	- Jira Ignore - will ignore the item on upload
	- Jira Update - will add the item using the update script

### Time Restriction 
While it can be adjusted there is a time range of 1 week for export.  The theory with that is that this script can be set as a scheduled task for the start of the week so that the users entries can automatically be added for the week.  Since its common for meetings to be added throughout the week, this time frame can also logically be changed to even a single day and set with the same Scheduled Task automation so that the user will freshly have their information daily to reduce edits.  

### Be Aware
If setting this script to run as a Scheduled Task, timing is important.  Since it would logically be run from the users machine it should be set for a time when the user will be logged on and Outlook open.  It is not a requirement for Outlook to be open for this script to work, however being open will ensure that the calendar is up to date.   

### What you are going to need...
**Profile ID**
	- Login to Jira and click to edit your profile
	- In the URL, after "/people/" this is the profile ID
**Jira Cloud Domain**
	- The main URL used for Jira
**API Token**
	- Instructions for creating API token can be found: [Jira Instructions] (https://support.atlassian.com/atlassian-account/docs/manage-api-tokens-for-your-atlassian-account/)
**Project Key**
	- Click on "View All Projects"
	- Filter for the Project you are working on
	- There is a field called "Key", this is the value you need
