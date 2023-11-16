This script will pull voice configurations from Microsoft Teams into an Excel workbook.

There are two dependencies.

- You need to be connected to your teams tenant (connect-microsoftteams) using a current version of the Microsoft Teams Powershell module 5.6.0. I am currently evaluating new versions of the powershell module.  I know it works with 5.3.  5.7.1 was recently released. 
- ImportExcel PowerShell module by Doug Finke (https://github.com/dfinke/ImportExcel) Current version - 7.8.6.  You will need to install the ImportExcel module by running <b>Install-Module -Name ImportExcel</b>.

The script will prompt you for a location to store the file. A directory will be created in that folder named "TeamsEnvironmentReports". The filename is automatically generated in the format of tenant ID-TeamsEnv-DateTimeStamp.xlsx (i.e. Contoso-TeamsEnv-04-10-2023.08.36.01.xlsx)

Once the script is launched you will be prompted if you want to include enterprise voice users or not.

The information that is pulled is: </p> 
<b>Green Tabs</b>:
Tenant information, PSTN Gateways, EV PSTN Usages, Voice Routes, Voice Routing Policies, Online Audio Conf. Routing, Dial Plan, Meeting QOS policies.

<b>Blue Tabs</b>:
Voice enabled users (this might take a while depending upon number of users), Phonenumbers, Dial in Conferencing Numbers, Auto Attendant, Voice App Policy, Call Queue, Resource accounts, caller ID policies, Calling Policies.

<b>Red tabs</b>:
Emergency Calling Policies, Emergency Call Routing Policies, Tenant Network Regions, Tenant Network Site Details, Tenant Network Subnet, Trusted IP address, LIS Locations, LIS Network, LIS WAP Information, LIS SWitch information, and LIS Port.

V3.3
Added Teams Audio Conferencing Policy on the EV Users Tab
Added licenses capabilities to the Phone numbers tab. 
Added tab for dial-in conferencing numbers that are not shared.  These should be numbers that are only used with in your tenant.  
