This script will pull voice configurations from Microsoft Teams into an Excel workbook.

There are two dependencies.

- You need to be connected to your teams tenant (connect-microsoftteams) using a current version of the Microsoft Teams Powershell module 5.1.0
- ImportExcel PowerShell module by Doug Finke (https://github.com/dfinke/ImportExcel) Current version - 7.8.4.  You will need to install the ImportExcel module by running <b>Install-Module -Name ImportExcel</b>.

The script will prompt you for a location to store the file. A directory will be created in that folder named "TeamsEnvironmentReports". The filename is automatically generated in the format of tenant ID-TeamsEnv-DateTimeStamp.xlsx (i.e. Contoso-TeamsEnv-04-10-2023.08.36.01.xlsx)

Once the script is launched you will be prompted if you want to include enterprise voice users or not.

The information that is pulled is Tenant information, PSTN Gateways, PSTN Usages, Voice Routes, Voice Routing Policies, Dial Plan, Voice enabled users (this might take a while depending upon number of users), Auto Attendant, Call Queue, Resource accounts, Meetign QOS policies, Emergency Calling Policies, Emergency Call Routing Policies, Tenant Network Site Details, Trusted IP address, LIS Locations, LIS Network Information, LIS WAP Information, LIS SWitch information, LIS Port, calling policies, caller ID policies, and Audio Conferencing policies.
