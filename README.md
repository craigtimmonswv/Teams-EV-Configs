This script will pull voice configuration from Microsoft Teams into an Excel workbook.

There are two dependencies.

- Connected to your teams tenant (connect-microsoftteams) using Microsoft Teams Powershell module 5.0.0
- ImportExcel PowerShell module by Doug Finke (https://github.com/dfinke/ImportExcel) Current version 7.8.4.  You will need to install the ImportExcel module by running <b>Install-Module -Name ImportExcel</b>.

It will prompt you for a location store the file. A directory will be created named "TeamsEnvironmentReports". The filename is automatically generated in the format of tenant ID-TeamsEnv-DateTimeStamp.xlsx (i.e. abcompany-TeamsEnv-03-24-2023.14.51.53.xlsx)

Once the script is launched you will be prompted if you want to include enterprise voice users or not.

The information that is pulled is Tenant information, PSTN Gateways, PSTN Usages, Voice Routes, Voice Routing Policies, Dial Plan, Voice enabled users - this might take a while depending upon number of users, Auto Attendant, Call Queue, Resource accounts, Meetign QOS policies, Emergency Calling Policies, Emergency Call Routing Policies, Tenant Network Site Details, Trusted IP address, LIS Locations, LIS Network Information, LIS WAP Information, LIS SWitch information, LIS Port, calling policies, caller ID policies, and Audio Conferencing policies.
