This script will pull voice configuration from Microsoft Teams into an Excel workbook.

There are two dependencies.

- Connected to your teams tenant (connect-microsoftteams)
- ImportExcel PowerShell module by Doug Finke (https://github.com/dfinke/ImportExcel)

It will prompt you for a location and file name to store the file and if you want to include enterprise voice users or not. 

The information that is pulled is 
PSTN Gateways,
PSTN Usages,
Voice Routes,
Voice Routing Policies,
Dial Plan,
Voice enabled users - this might take a while depending upon number of users,
Emergency Calling Policies,
Emergency Call Routing Policies,
Tenant Network Site Details,
LIS Locations,
LIS Network Information,
LIS WAP Information,
LIS SWitch information,
LIS Port,
Auto Attendant,
Call Queue.
