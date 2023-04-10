<#
You will need have the "ImportExcel" Module installed for this to properly run. 
You can get it here:
https://www.powershellgallery.com/packages/ImportExcel/7.8.4
To install it run: 
Install-Module -Name ImportExcel -RequiredVersion 7.8.4
Import-Module -Name ImportExcel
This script will pull the basic voice environment from the Teams tenant. Items gathered:
Infrastructure items and makes the tabs green:
    Tenant Information
    PSTN Gateways
    PSTN Usages
    Voice Routes
    Voice Routing Policies
    Dial Plans
    Teams Meeting Settings - If QOS is enabled, it will make the cell H2 green.  If it isn't enabled, it will make cell H2 red.

User details, and various policies.  It will make the tabs blue.  Items gathered:
    Voice enabled users - this might take a while depending upon number of users.  It is optional.
        Displayname, UPN, City, State, Country, Usage Location, Lineuri, Licenses, Dial Plan, Voice routing policy, 
        Enterprise voice enabled, Teams upgrade policy, teams effective mode, emergency calling policy, emergency call routing policy, 
        Teams calling policy, Teams meeting policy, and Audio Conferencing Policy
    Auto-Attendant details
    Call Queue details
    Resource account details
    Caller ID Policy
    Calling Policies
    Audio Conferencing policies
Emergency services items
    Emergency Calling Policies
    Emergency Call Routing Policies
    Tenant Network Region Details
    Tenant Network Site Details
    Trusted IP Address
    LIS Locations
    LIS Network Information (Subnets)
    LIS WAP Information
    LIS SWitch information
    LIS Port information

You will be prompted to enter a location to store the spreadsheet. This will be directory and should be formated like "C:\scriptoutput".  
It will then create a folder called "TeamsEnvironmentReports".  This folder will hold the output of the spreadsheet and any error logs.  

The spreadsheet will have a name that contains the tenant, and date/time stamp (format "Contoso-TeamsEnv-11-18-2022.12.49.11.xlsx".)
A few changes have been made:
    1. Format of the voice routing policy will have the OnlinePstnUages on one line, separted by commas.  
    2. Added Error checking.  A log file will be created with the name similar to the file name, but will 
       include errorlog in the filename.
    3. Added various policy reports (Meeting, Calling Policies, Caller ID Policy, Application Permission Policy, 
       Teams Meeting Configuration, Audio Conferencing)
    4. Added feature types (licenses), assigned plans, and various Teams policies to the EV users report.  Some of these 
       will be empty unless the user is using calling plans, operator connect, DRaaS, or doing something with Video Interop.
    5. Added Network Region Information
    6. Added coloring of tabs based upon the type of information (Green = Infrastructre, Blue = User Details, Red = Emergency LIS)
    7. Created functions for each tab being created.   
#>
$date = get-date -Format "MM/dd/yyyy HH:mm"
$tabFreeze = "PSTN Gateways","EV Users","Voice Routes","Call Queue","LIS Location","Calling Policies","Auto Attendant"

Function Write-DataToExcel
    {
        param ($filelocation, $details, $tabname, $tabcolor)
        $excelpackage = Open-ExcelPackage -Path $filelocation 
        $ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName $tabname 
        $ws.Workbook.Worksheets[$ws.index].TabColor = $tabcolor
        if ($tabname -eq "Meeting-QOS Policies")
            {Add-ConditionalFormatting -Worksheet $ws -Address "H2:H2" -RuleType ContainsText -ConditionValue "FALSE" -backGroundColor red -ForegroundColor white
            Add-ConditionalFormatting -Worksheet $ws -Address "H2:H2" -RuleType ContainsText -ConditionValue "TRUE" -backGroundColor Green -ForegroundColor white
            }
    
        if ($tabFreeze.Contains($tabname))
            {$details | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter -FreezePane 2,3}
        else 
            {$details | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter -FreezeTopRow       }
        Clear-Variable details 
        Clear-Variable filelocation
        Clear-Variable tabname
        Clear-Variable TabColor
    }
Function write-Errorlog
{
    param ($logfile, $errordata, $msgData)
    $errordetail = '"'+ $date + '","' + $msgData + '","' + $errordata + '"'
    Write-Host $errordetail
    $errordetail |  Out-File -FilePath $logname -Append 
    Clear-Variable errordetail, msgData
}
Function Write-TenantInfo
{
    Write-host "Getting Tenant Information"
    $tenatDetail = Get-CsTenant
    $detail = New-Object PSObject
    $detail | add-Member -MemberType NoteProperty -Name "DisplayName" -Value $tenatDetail.DisplayName
    $detail | add-Member -MemberType NoteProperty -Name "TeamsUpgradeEffectiveMode" -Value $tenatDetail.TeamsUpgradeEffectiveMode
    $detail | add-Member -MemberType NoteProperty -Name "TenantId" -Value $tenatDetail.TenantId
    $Detail |Export-Excel -Path $filelocation -WorksheetName "Tenant info" -AutoFilter -AutoSize
    $excel = Open-ExcelPackage -Path $filelocation 
    $Green = "Green"
    $Green = [System.Drawing.Color]::$green 
    $excel.Workbook.Worksheets[1].TabColor = $Green  
    Close-ExcelPackage -ExcelPackage $excel
}
Function Write-PSTNGateways
{
    # Extract PSTN Gateways
    Write-Host 'Getting Online PSTN Gateway Details'
    $Details = @()
    Try { $PSTNGWs = Get-CsOnlinePSTNGateway -ErrorAction Stop -WarningAction SilentlyContinue}
    Catch 
        {
            $msgdata = "Error getting PSTN Gateway Details."
            write-Errorlog $logfile $error[0].exception.message $msgData
            Clear-Variable msgData
        }
    if ($PSTNGWs.count -ne 0)
        {
            foreach ($GW in $PSTNGWs)
            {       
                $detail = New-Object PSObject
                $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $GW.Identity
                $detail | add-Member -MemberType NoteProperty -Name "Fqdn" -Value $GW.Fqdn
                $detail | Add-Member -MemberType NoteProperty -Name "SIPSignalingPort" -Value $GW.SipSignalingPort
                $detail | Add-Member -MemberType NoteProperty -Name "FailoverTimeSeconds" -Value $GW.FailoverTimeSeconds
                $detail | Add-Member -MemberType NoteProperty -Name "ForwardCallHistory" -Value $GW.ForwardCallHistory
                $detail | Add-Member -MemberType NoteProperty -Name "ForwardPai" -Value $GW.ForwardPai
                $detail | Add-Member -MemberType NoteProperty -Name "SendSipOptions" -Value $GW.SendSipOptions
                $detail | Add-Member -MemberType NoteProperty -Name "MaxConcurrentSessions" -Value $GW.MaxConcurrentSessions
                $detail | Add-Member -MemberType NoteProperty -Name "Enabled" -Value $GW.Enabled
                $detail | Add-Member -MemberType NoteProperty -Name "BypassMode" -Value $GW.BypassMode
                $detail | Add-Member -MemberType NoteProperty -Name "MediaBypass" -Value $GW.MediaBypass
                $detail | Add-Member -MemberType NoteProperty -Name "GatewaySiteId" -Value $GW.GatewaySiteId
                $detail | Add-Member -MemberType NoteProperty -Name "PidfLoSupported" -Value $GW.PidfLoSupported
                $detail | Add-Member -MemberType NoteProperty -Name "ProxySbc" -Value $GW.ProxySbc
                $detail | Add-Member -MemberType NoteProperty -Name "GatewaySiteLbrEnabled" -Value $GW.GatewaySiteLbrEnabled
                $detail | Add-Member -MemberType NoteProperty -Name "FailoverResponseCodes" -Value $GW.FailoverResponseCodes.Replace(",",", ")
                $Details += $detail
            }
        }
        Else {$details = "No Data to Display"}
        $tabname = 'PSTN Gateways'
        $tabcolor = "Green"
        Write-DataToExcel $filelocation $details $tabname $tabcolor
}
Function write-PSTNUsages
{
    # Get PSTN Usages
    Write-Host 'Getting PSTN Usages'
    try {$PSTNUSAGEs = Get-CsOnlinePstnUsage -ErrorAction Stop }
    Catch 
        {
            $msgdata = "Error getting PSTN Usage Details."
            write-Errorlog $logfile $error[0].exception.message $msgData
            Clear-Variable msgData
        }
    if ($PSTNUSAGEs.count -ne 0)
    {
        $details =@()
        foreach ($PSTNUsage in $PSTNUSAGEs)
        {   
            foreach ($u in $PSTNUSAGE.Usage)
            {
                $detail = New-Object PSObject
                $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $PSTNUSAGE.Identity
                $detail | add-Member -MemberType NoteProperty -Name "Usage" -Value $u
                $details += $detail
            }
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = 'EV PSTN Usages'
    $tabcolor = 'Green'
    Write-DataToExcel $filelocation  $details $tabname $tabcolor
}
Function Write-VoiceRoutes
{
    # Get Voice Routes
    Write-Host 'Getting Voice Routes'

    try {$VRs = Get-CsOnlineVoiceRoute -ErrorAction stop}
    Catch 
        {
            $msgdata = "Error getting Voice Routes Details."
            write-Errorlog $logfile $error[0].exception.message $msgData
            Clear-Variable msgData
        }
    if ($vrs.count -ne 0)
    {
        $Details = @()
        foreach ($VR in $VRs)
        {   
            [string] $usage= $vr.OnlinePstnUsages
            [string] $pstngw =$vr.OnlinePstnGatewayList
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $vr.Name
            $detail | Add-Member -MemberType NoteProperty -Name "NumberPattern" -Value $vr.NumberPattern
            $detail | Add-Member -MemberType NoteProperty -Name "OnlinePstnUsages" -Value $usage
            $detail | Add-Member -MemberType NoteProperty -Name "OnlinePstnGatewayList " -Value $pstngw.Replace(" ",", ") 
            $details += $detail
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = 'Voice Routes'
    $tabcolor = 'Green'
    Write-DataToExcel $filelocation  $details $tabname $tabcolor
}

Function Write-VoiceRoutingPolicies
{
    # Get Voice Routing Policies
    Write-Host 'Getting Voice Routing Policies'
    try {$vrps = Get-CsOnlineVoiceRoutingPolicy -ErrorAction:Stop}
    Catch 
        {
            $msgdata = "Error getting Voice Routing Policy Details."
            write-Errorlog $logfile $error[0].exception.message $msgData
            Clear-Variable msgData
        } 
    IF ($vrps.count -ne 0)
    {
        $details = @()
        foreach ($VRP in $VRPs)
        {   
            try {[string] $opu = (Get-CsOnlineVoiceRoutingPolicy -Identity $vrp.Identity -erroraction Stop| Select-Object OnlinePstnUsages).OnlinePstnUsages }
            catch 
            {
                $msgdata = "Error getting Voice Routing Policy Details."
                write-Errorlog $logfile $error[0].exception.message $msgData
                Clear-Variable msgData
            }
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $VRP.Identity
            $detail | add-Member -MemberType NoteProperty -Name "OnlinePstnUsages" -Value $opu.Replace(" ",", ")
            $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $VRP.Description
            $detail | add-Member -MemberType NoteProperty -Name "RouteType" -Value $VRP.RouteType
            $details += $detail
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "Voice Routing Policies"
    $tabcolor = 'Green'
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-DialPlans
{
    # Get Dial Plan info
    Write-Host 'Getting Dial Plan Details'

    try {$DPs=Get-CsTenantDialPlan -ErrorAction Stop}
    catch 
        {
            $msgdata = "Error getting Dial Plans Details."
            write-Errorlog $logfile $error[0].exception.message $msgData
            Clear-Variable msgData
        }
    if ($DPs.count -ne 0)
    {
        $Details = @()
        foreach ($dp in $DPs)
        {   
        foreach ($rule in $dp.NormalizationRules)
            {
                # Creating an array to store the variables from the dial plans. 
                $detail = New-Object PSObject
                $detail | add-Member -MemberType NoteProperty -Name "Parent" -Value $dp.Identity.remove(0,4)
                $detail | Add-Member -MemberType NoteProperty -Name "Description" -Value $rule.Description
                $detail | Add-Member -MemberType NoteProperty -Name "Name" -Value $rule.Name
                $detail | Add-Member -MemberType NoteProperty -Name "Pattern" -Value $rule.Pattern
                $detail | Add-Member -MemberType NoteProperty -Name "Translation" -Value $rule.Translation
                $detail | Add-Member -MemberType NoteProperty -Name "IsInternalExtension" -Value $rule.IsInternalExtension
                # Adding array from one dial plan to an array with all the dial plans. 
                $Details += $detail
            }
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "Dial Plan"
    $tabcolor = "Green"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-TeamsMeetingsSettings
{
    # Teams Meeting Settings
    Write-Host 'Getting Teams Meeting Configuration Details'
    $MTGConfigs = Get-CsTeamsMeetingConfiguration
    if ($MTGConfigs.count -ne 0)
    {
        $details = @()
        $detail = New-Object PSObject
        $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $MTGConfigs.Identity
        $detail | Add-Member -MemberType NoteProperty -Name "LogoURL" -Value $MTGConfigs.LogoURL
        $detail | Add-Member -MemberType NoteProperty -Name "LegalURL" -Value $MTGConfigs.LegalURL
        $detail | Add-Member -MemberType NoteProperty -Name "HelpURL" -Value $MTGConfigs.HelpURL
        $detail | Add-Member -MemberType NoteProperty -Name "CustomFooterText" -Value $MTGConfigs.CustomFooterText
        $detail | Add-Member -MemberType NoteProperty -Name "DisableAnonymousJoin" -Value $MTGConfigs.DisableAnonymousJoin
        $detail | Add-Member -MemberType NoteProperty -Name "DisableAppInteractionForAnonymousUsers" -Value $MTGConfigs.DisableAppInteractionForAnonymousUsers
        $detail | Add-Member -MemberType NoteProperty -Name "EnableQoS" -Value $MTGConfigs.EnableQoS
        $detail | Add-Member -MemberType NoteProperty -Name "ClientAudioPort" -Value $MTGConfigs.ClientAudioPort
        $detail | Add-Member -MemberType NoteProperty -Name "ClientAudioPortRange" -Value $MTGConfigs.ClientAudioPortRange
        $detail | Add-Member -MemberType NoteProperty -Name "ClientVideoPort" -Value $MTGConfigs.ClientVideoPort
        $detail | Add-Member -MemberType NoteProperty -Name "ClientVideoPortRange" -Value $MTGConfigs.ClientVideoPortRange
        $detail | Add-Member -MemberType NoteProperty -Name "ClientAppSharingPort" -Value $MTGConfigs.ClientAppSharingPort
        $detail | Add-Member -MemberType NoteProperty -Name "ClientAppSharingPortRange" -Value $MTGConfigs.ClientAppSharingPortRange
        $detail | Add-Member -MemberType NoteProperty -Name "ClientMediaPortRangeEnabled" -Value $MTGConfigs.ClientMediaPortRangeEnabled
        $details += $detail
    }
    Else {$details = "No Data to Display"}
    $tabname = "Meeting-QOS Policies"
    $tabcolor = 'Green'
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-EVUsers
{
    Write-Host 'Getting Voice Enabled Users'
    $Details = @()
    try {$users =  Get-CsOnlineUser -Filter {EnterpriseVoiceEnabled -eq $true} -ErrorAction Stop }
    catch 
    {
        $msgdata = "Error getting Enterprise Voice User Details."
        write-Errorlog $logfile $error[0].exception.message $msgData
        Clear-Variable msgData
    }
    if ($users.Count -ne 0)
    {
        $details = @()
        foreach ($user in $users)
        {
            [string]$license = $user.featuretypes
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Displayname" -Value $user.displayname
            $detail | add-Member -MemberType NoteProperty -Name "UPN" -Value $user.UserPrincipalName
            $detail | add-Member -MemberType NoteProperty -Name "City" -Value $user.City
            $detail | add-Member -MemberType NoteProperty -Name "State" -Value $user.StateOrProvince
            $detail | add-Member -MemberType NoteProperty -Name "Country" -Value $user.Country
            $detail | add-Member -MemberType NoteProperty -Name "UsageLocation" -Value $user.UsageLocation
            $detail | add-Member -MemberType NoteProperty -Name "Lineuri" -Value $user.LineUri
            $detail | add-Member -MemberType NoteProperty -Name "Licenses" -Value $license.Replace(" ",", ")
            $detail | add-Member -MemberType NoteProperty -Name "Dial Plan" -Value $user.TenantDialPlan
            $detail | add-Member -MemberType NoteProperty -Name "Voice Routing Policy" -Value $user.OnlineVoiceRoutingPolicy
            $detail | add-Member -MemberType NoteProperty -Name "EV Enabled" -Value $user.EnterpriseVoiceEnabled
            $detail | add-Member -MemberType NoteProperty -Name "Teams Upgrade Policy" -Value $user.TeamsUpgradePolicy
            $detail | add-Member -MemberType NoteProperty -Name "Teams Effective Mode" -Value $user.TeamsUpgradeEffectiveMode
            $detail | add-Member -MemberType NoteProperty -Name "Emergency Calling Policy" -Value $user.TeamsEmergencyCallingPolicy 
            $detail | add-Member -MemberType NoteProperty -Name "Emergency Call Routing Policy" -Value $user.TeamsEmergencyCallRoutingPolicy  
            #$detail | add-Member -MemberType NoteProperty -Name "Teams Carrier Emergency Call Routing Policy" -Value $user.TeamsCarrierEmergencyCallRoutingPolicy
            $detail | add-Member -MemberType NoteProperty -Name "Teams Calling Policy" -Value $user.TeamsCallingPolicy
            $detail | add-Member -MemberType NoteProperty -Name "Teams Meeting Policy" -Value $user.TeamsMeetingPolicy
            $detail | Add-Member -MemberType NoteProperty -Name "Audio Conferencing Policy" -Value $user.OnlineAudioConferencingRoutingPolicy
            $details += $detail
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "EV Users"
    $tabcolor = 'Blue'
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-AutoAttendants
{
    #Get Auto Attendant Details
    Write-Host "Getting Auto Attendant Details"
    $AAs = Get-CsAutoAttendant
    $details =@()
    foreach ($aa in $AAs)
    {
        foreach ($RA in $aa.ApplicationInstances)
        {
            Try {$ResouceAct = Get-CsOnlineApplicationInstance -Identity $ra -ErrorAction stop}
            catch
            {
                $msgdata = "Error getting AA Application instance."
                write-Errorlog $logfile $error[0].exception.message $msgData
                Clear-Variable msgData
            }
        }
        try { $operatorID = ((Get-CsAutoAttendant -NameFilter $aa.Name | Select-Object operator).operator).id }
        catch
            {
                $msgdata = "Error getting AA Application instance."
                write-Errorlog $logfile $error[0].exception.message $msgData
                Clear-Variable msgData
            }
        if (!($operatorID))
                {$operator = "No Operator Defined"}
        Else
                {
                    try {$Operator = (Get-CsOnlineUser -Identity $operatorID -erroraction Stop | Select-Object UserPrincipalName).UserPrincipalName}
                    catch 
                    {
                        $msgdata = "Error getting AA Application instance."
                        write-Errorlog $logfile $error[0].exception.message $msgData
                        Clear-Variable msgData
                    }
                }
            $detail = New-Object PSObject
            $detail | Add-Member NoteProperty -Name "AAName" -Value $aa.name
            $detail | Add-Member NoteProperty -Name "Identity" -Value $aa.identity
            $detail | Add-Member NoteProperty -Name "Language" -Value $aa.LanguageId
            $detail | Add-Member NoteProperty -Name "TimeZone" -Value $aa.timezoneid
            $detail | Add-Member NoteProperty -Name "Operator" -Value $Operator
            $detail | Add-Member NoteProperty -Name "VoiceResponseEnabled" -Value $aa.VoiceresponseEnabled
            $detail | Add-Member NoteProperty -Name "ResourceAccount" -Value $ResouceAct.UserPrincipalName
            $detail | Add-Member NoteProperty -Name "Phone Number" -Value $ResouceAct.PhoneNumber
            $details += $detail
            Clear-Variable detail
    }
    $tabname = "Auto Attendant"
    $tabcolor = 'Blue'
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-CallQueues
{
    # Get Call Queues Details
    Write-Host "Getting Call Queue Details"
    $CQs = Get-CsCallQueue -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    if ($CQs.count -ne 0)
    {
        $Details = @()
        foreach ($CQ in $CQs)
        {
            $detail = New-Object PSObject
            $detail | Add-Member NoteProperty -Name "CQName" -Value $CQ.name
            $detail | Add-Member NoteProperty -Name "Identity" -Value $CQ.identity
            $detail | Add-Member NoteProperty -Name "RoutingMethod" -Value $CQ.RoutingMethod
            $detail | Add-Member NoteProperty -Name "AllowOptOut" -Value $CQ.AllowOptOut                
            $detail | Add-Member NoteProperty -Name "ConferenceMode" -Value $CQ.ConferenceMode
            $detail | Add-Member NoteProperty -Name "PresenceBasedRouting" -Value $CQ.PresenceBasedRouting
            foreach ($a in $cq.Agents.objectid)
                {
                    try {$agent=(get-csonlineuser -Identity $a -erroraction SilentlyContinue| Select-Object UserPrincipalName).UserPrincipalName + ","}
                    Catch 

                    {
                        $msgdata = "Error getting CQ Agents."
                        write-Errorlog $logfile $error[0].exception.message $msgData
                        Clear-Variable msgData
                    }
                    $agents +=$agent
                }
            $detail | Add-Member NoteProperty -Name "Agents" -Value $agents
            $detail | Add-Member NoteProperty -Name "AgentAlertTime" -Value $CQ.AgentAlertTime
            $detail | Add-Member NoteProperty -Name "LanguageId" -Value $CQ.LanguageId
            $detail | Add-Member NoteProperty -Name "OverflowThreshold" -Value $CQ.OverflowThreshold
            $detail | Add-Member NoteProperty -Name "OverflowAction" -Value $CQ.OverflowAction
            try {$OFATarget = ((Get-CsCallQueue -NameFilter $cq.Name| Select-Object OverflowActionTarget).OverflowActionTarget).id }
            catch 
                {
                    $msgdata = "Error getting CQ Over Flow Action Targets."
                    write-Errorlog $logfile $error[0].exception.message $msgData
                    Clear-Variable msgData
                }
                if ($OFATarget)
                    {
                        try { $OFATargetUser = (get-csonlineuser -Identity $OFATarget -erroraction SilentlyContinue| Select-Object UserPrincipalName).UserPrincipalName}
                        catch 
                        {
                            $msgdata = "Error getting CQ Over Flow Action Targets."
                            write-Errorlog $logfile $error[0].exception.message $msgData
                            Clear-Variable msgData
                        }
                    }
            $detail | Add-Member NoteProperty -Name "OverflowActionTarget" -Value $OFATargetUser
            $detail | Add-Member NoteProperty -Name "OverflowSharedVoicemailTextToSpeechPrompt" -Value $CQ.OverflowSharedVoicemailTextToSpeechPrompt
            $detail | Add-Member NoteProperty -Name "EnableOverflowSharedVoicemailTranscription" -Value $CQ.EnableOverflowSharedVoicemailTranscription
            $detail | Add-Member NoteProperty -Name "TimeoutThreshold" -Value $CQ.TimeoutThreshold
            $detail | Add-Member NoteProperty -Name "TimeoutAction" -Value $CQ.TimeoutAction
            try {     $TOATarget = ((Get-CsCallQueue -NameFilter $cq.Name| Select-Object TimeoutActionTarget).TimeoutActionTarget).id}
            catch 
            {
                $msgdata = "Error getting CQ Timeout Action Targets."
                write-Errorlog $logfile $error[0].exception.message $msgData
                Clear-Variable msgData
            }
            if ($TOATarget)
                {
                    try {$TOATargettUser = (get-csonlineuser -Identity $TOATarget -erroraction SilentlyContinue | Select-Object UserPrincipalName).UserPrincipalName}
                    catch 
                    {
                        $msgdata = "Error getting CQ Timeout Action Targets."
                        write-Errorlog $logfile $error[0].exception.message $msgData 
                        Clear-Variable msgData
                    }
                }
            $detail | Add-Member NoteProperty -Name "TimeoutActionTarget" -Value $TOATargettUser
            $detail | Add-Member NoteProperty -Name "TimeoutSharedVoicemailTextToSpeechPrompt" -Value $CQ.TimeoutSharedVoicemailTextToSpeechPrompt
            $detail | Add-Member NoteProperty -Name "EnableTimeoutSharedVoicemailTranscription" -Value $CQ.EnableTimeoutSharedVoicemailTranscription
            $details += $detail
            try {Clear-Variable agent -ErrorAction SilentlyContinue}
            Catch{}
            try {Clear-Variable agents -ErrorAction SilentlyContinue}
            Catch{}
            try {Clear-Variable TOATarget -ErrorAction SilentlyContinue}
            Catch{}
            try {Clear-Variable OFATarget -ErrorAction SilentlyContinue} 
            Catch{}      
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "Call Queue"
    $tabcolor = 'Blue'
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}
Function Write-ResourceAccounts
{
    #Get Resource Account 
    Write-Host 'Getting Resource Account Information'
    $RAs = get-csonlineapplicationInstance
    if ($RAs -ne 0)
    {
        $details = @()
        foreach ($ra in $RAs)
        {
            $detail = New-Object PSObject
            $detail | Add-Member NoteProperty -Name "DisplayName" -Value $ra.DisplayName
            $detail | Add-Member NoteProperty -Name "UserPrincipalName" -Value $ra.UserPrincipalName
            $detail | Add-Member NoteProperty -Name "ObjectId" -Value $ra.ObjectId
            $detail | Add-Member NoteProperty -Name "PhoneNumber" -Value $ra.PhoneNumber
            $detail | Add-Member NoteProperty -Name "ApplicationId" -Value $ra.ApplicationId
            if ($ra.ApplicationId -eq '11cd3e2e-fccb-42ad-ad00-878b93575e07')
            {$AppType = "Call Queue"}
            else { $AppType = "Auto-Attendant"}
            $detail | Add-Member NoteProperty -Name "Application Type" -Value $AppType
            $Details += $detail
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "Res Account Details"
    $tabcolor = 'Blue'
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}
Function Write-CallerIDPolicy
{
    # Caller ID Policy
    Write-Host 'Getting Caller ID Policy Details'
    $CIDPs = Get-CsCallingLineIdentity
    If ($cidps.count -ne 0)
    {
        $details = @()
        foreach ($cidp in $cidps)
        {
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $cidp.Identity
            $detail | Add-Member -MemberType NoteProperty -Name "Description" -Value $cidp.Description
            $detail | Add-Member -MemberType NoteProperty -Name "EnableUserOverride" -Value $cidp.EnableUserOverride
            $detail | Add-Member -MemberType NoteProperty -Name "ServiceNumber" -Value $cidp.ServiceNumber
            $detail | Add-Member -MemberType NoteProperty -Name "CallingIDSubstitute" -Value $cidp.CallingIDSubstitute
            $detail | Add-Member -MemberType NoteProperty -Name "BlockIncomingPstnCallerID" -Value $cidp.BlockIncomingPstnCallerID
            $detail | Add-Member -MemberType NoteProperty -Name "ResourceAccount" -Value $cidp.ResourceAccount
            $detail | Add-Member -MemberType NoteProperty -Name "CompanyName" -Value $cidp.CompanyName
            $details += $detail
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "Caller ID Policies"
    $tabcolor = "Blue"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}
Function Write-CallingPolicies
{
    # Get Calling Policies
    Write-Host 'Getting Calling Policies Details'
    $CPs = Get-CsTeamsCallingPolicy
    If ($cps.count -ne 0)
    {
        $Details = @()
        foreach ($cp in $CPS)
        {
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $cp.Identity.remove(0,4)
            $detail | Add-Member -MemberType NoteProperty -Name "Description" -Value $cp.Description
            $detail | Add-Member -MemberType NoteProperty -Name "AllowPrivateCalling" -Value $cp.AllowPrivateCalling
            $detail | Add-Member -MemberType NoteProperty -Name "AllowWebPSTNCalling" -Value $cp.AllowWebPSTNCalling
            $detail | Add-Member -MemberType NoteProperty -Name "AllowSIPDevicesCalling" -Value $cp.AllowSIPDevicesCalling
            $detail | Add-Member -MemberType NoteProperty -Name "AllowVoicemail" -Value $cp.AllowVoicemail
            $detail | Add-Member -MemberType NoteProperty -Name "AllowCallGroups" -Value $cp.AllowCallGroups
            $detail | Add-Member -MemberType NoteProperty -Name "AllowDelegation" -Value $cp.AllowDelegation
            $detail | Add-Member -MemberType NoteProperty -Name "AllowCallForwardingToUser" -Value $cp.AllowCallForwardingToUser
            $detail | Add-Member -MemberType NoteProperty -Name "AllowCallForwardingToPhone" -Value $cp.AllowCallForwardingToPhone
            $detail | Add-Member -MemberType NoteProperty -Name "PreventTollBypass" -Value $cp.PreventTollBypass
            $detail | Add-Member -MemberType NoteProperty -Name "BusyOnBusyEnabledType" -Value $cp.BusyOnBusyEnabledType
            $detail | Add-Member -MemberType NoteProperty -Name "MusicOnHoldEnabledType" -Value $cp.MusicOnHoldEnabledType
            $detail | Add-Member -MemberType NoteProperty -Name "AllowCloudRecordingForCalls" -Value $cp.AllowCloudRecordingForCalls
            $detail | Add-Member -MemberType NoteProperty -Name "AllowTranscriptionForCalling" -Value $cp.AllowTranscriptionForCalling
            $detail | Add-Member -MemberType NoteProperty -Name "PopoutForIncomingPstnCalls" -Value $cp.PopoutForIncomingPstnCalls
            $detail | Add-Member -MemberType NoteProperty -Name "PopoutAppPathForIncomingPstnCalls" -Value $cp.PopoutAppPathForIncomingPstnCalls
            $detail | Add-Member -MemberType NoteProperty -Name "LiveCaptionsEnabledTypeForCalling" -Value $cp.LiveCaptionsEnabledTypeForCalling
            $detail | Add-Member -MemberType NoteProperty -Name "AutoAnswerEnabledType" -Value $cp.AutoAnswerEnabledType
            $detail | Add-Member -MemberType NoteProperty -Name "SpamFilteringEnabledType" -Value $cp.SpamFilteringEnabledType
            $detail | Add-Member -MemberType NoteProperty -Name "CallRecordingExpirationDays" -Value $cp.CallRecordingExpirationDays
            $detail | Add-Member -MemberType NoteProperty -Name "AllowCallRedirect" -Value $cp.AllowCallRedirect
            $details += $detail
        }
    }   
    Else {$details = "No Data to Display"}
    $tabname = "Calling Policies"
    $tabcolor = "Blue"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}
Function Write-AudioConferencingPolicy
{
    # Audio Conferencing Settings
    Write-Host 'Getting Audio Conferencing Details'
    $AudConfs = Get-CsTeamsAudioConferencingPolicy 
    if ($AudConfs.count -ne 0)
    {
        $details = @()
        foreach ($AudConf in $AudConfs)
        { 
            [string]$dialin = $AudConf.MeetingInvitePhoneNumbers

            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $AudConf.Identity
            $detail | Add-Member -MemberType NoteProperty -Name "MeetingInvitePhoneNumbers" -Value $dialin.Replace(" ",", ")
            $detail | Add-Member -MemberType NoteProperty -Name "AllowTollFreeDialin" -Value $AudConf.AllowTollFreeDialin
            $details += $detail
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "Audio Conferencing Policies"
    $tabcolor = "Blue"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-EmergencyCallingPolicy
{
    # Get Emergency Calling Policies
    Write-Host 'Getting Emergency Calling Policies'
    $Details = @()
    try {$ercallpolicies = Get-CsTeamsEmergencyCallingPolicy -ErrorAction Stop }
    catch 
    {
        $msgdata = "Error getting Emergency Calling Policy Details."
        write-Errorlog $logfile $error[0].exception.message $msgData
        Clear-Variable msgData
    }
    if ($ercallpolicies -ne 0)
    {
        foreach ($ercp in $ercallpolicies)
        {
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $ercp.Identity
            $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $ercp.Description
            $detail | add-Member -MemberType NoteProperty -Name "NotificationGroup" -Value $ercp.NotificationGroup
            $detail | add-Member -MemberType NoteProperty -Name "ExternalLocationLookupMode" -Value $ercp.ExternalLocationLookupMode
            $detail | add-Member -MemberType NoteProperty -Name "NotificationDialOutNumber" -Value $ercp.NotificationDialOutNumber
            $detail | add-Member -MemberType NoteProperty -Name "NotificationMode" -Value $ercp.NotificationMode
            $details += $detail  
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "Emergency Calling Policies"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-EmergencyCallRouting
{
    # Get Emergency Call Routing Policy
    Write-Host 'Getting Emergency Call Routing Policies'
    $Details = @()
    try {$ecrps = Get-CsTeamsEmergencyCallRoutingPolicy -ErrorAction Stop }
    catch 
    {
        $msgdata = "Error getting Emergency Call Routing Policy Details."
        write-Errorlog $logfile $error[0].exception.message $msgData
        Clear-Variable msgData
    }
    if ($ecrps.count -ne 0)
    {
        foreach ($ecrp in $ecrps)
            {
                $numbers = Get-CsTeamsEmergencyCallRoutingPolicy -Identity $ecrp.identity
                foreach ($number in $numbers.EmergencyNumbers)
                    {
                        $detail = New-Object PSObject
                        $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $ecrp.Identity
                        $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $ecrp.Description
                        $detail | add-Member -MemberType NoteProperty -Name "emergencydialstring" -Value $number.emergencydialstring
                        $detail | add-Member -MemberType NoteProperty -Name "EmergencyDialMask" -Value $number.emergencydialmask
                        $detail | add-Member -MemberType NoteProperty -Name "OnlinePSTNUsage" -Value $number.OnlinePSTNUsage
                        $detail | add-Member -MemberType NoteProperty -Name "AllowEnhancedEmergencyServices" -Value $ecrp.AllowEnhancedEmergencyServices
                        $details  += $detail  
                    }
            }
    }
    Else {$details = "No Data to Display"}
    $tabname = "Emergency Call Routing Policies"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-NetworkSiteDetails
{
    # Get Tenant Network Site Details
    Write-Host 'Getting Tenant Network Site Details'
    $Details = @()
    try {$sites = Get-CsTenantNetworkSite -ErrorAction Stop}
    catch 
        {
            $msgdata = "Error getting Tenant Network Site Details."
            write-Errorlog $logfile $error[0].exception.message $msgData
            Clear-Variable msgData
        }
    if ($sites.count -ge 1)
        {
            foreach ($site in $sites)
            {
                $detail = New-Object PSObject
                $detail | add-Member -MemberType NoteProperty -Name "Subnets" -Value $site.Subnets
                $detail | add-Member -MemberType NoteProperty -Name "Postalcodes" -Value $site.Postalcodes
                $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $site.Identity
                $detail | add-Member -MemberType NoteProperty -Name "NetworkSiteID" -Value $site.NetworkSiteID
                $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $site.Description
                $detail | add-Member -MemberType NoteProperty -Name "NetworkRegionID" -Value $site.NetworkRegionID
                $detail | add-Member -MemberType NoteProperty -Name "LocationPolicy" -Value $site.LocationPolicy
                $detail | add-Member -MemberType NoteProperty -Name "EnableLocationBasedRouting" -Value $site.EnableLocationBasedRouting
                $detail | add-Member -MemberType NoteProperty -Name "SiteAddress" -Value $site.SiteAddress
                $detail | add-Member -MemberType NoteProperty -Name "EmergencyCallRoutingPolicy" -Value $site.EmergencyCallRoutingPolicy
                $detail | add-Member -MemberType NoteProperty -Name "EmergencyCallingPolicy" -Value $site.EmergencyCallingPolicy
                $detail | add-Member -MemberType NoteProperty -Name "NetworkRoamingPolicy" -Value $site.NetworkRoamingPolicy
                $details += $detail  
            }
        }
    
    Else {$details = "No Data to Display"}
    $tabname = "Tenant Network Site Details"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}
Function Write-NetworkRegion
{   
    Write-Host "Getting Tenant Network Region"
    $Details = @()
    $regions = Get-CsTenantNetworkRegion
    if ($regions.count -ge 1)
    {
        foreach ($region in $regions)
        {
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $region.Identity
            $detail | add-Member -MemberType NoteProperty -Name "NetworkRegionID" -Value $region.NetworkRegionID
            $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $region.Description
            $detail | add-Member -MemberType NoteProperty -Name "CentralSite" -Value $region.CentralSite
            $Details += $detail
        }
    }
    else {$details = "No Data to Display"}
    $tabname = "Tenant Network Region"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-NetworkSubnetDetails
{   
    Write-Host "Getting Tenant Network Subnets"
    $Details = @()
    $subnets = Get-CsTenantNetworkSubnet
    if ($subnets.count -ge 1)
    {
        foreach ($subnet in $subnets)
        {
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $subnets.Identity
            $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $subnets.Description
            $detail | add-Member -MemberType NoteProperty -Name "NetworkSiteID" -Value $subnets.NetworkSiteID
            $detail | add-Member -MemberType NoteProperty -Name "SubnetID" -Value $subnets.SubnetID
            $detail | add-Member -MemberType NoteProperty -Name "MaskBits" -Value $subnets.MaskBits
            $Details += $detail
        }
    }
    else {$details = "No Data to Display"}
    $tabname = "Tenant Network Subnet"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-TrustedIPs
{
    # Get Tenant Trusted IP Addresses
    Write-Host 'Getting Tenant Trusted IP Addresses'
    $Details = @()
    try {$TrustedIPs = get-CsTenantTrustedIPAddress -ErrorAction Stop}
    catch 
        {
            $msgdata = "Error getting Trusted IP Address Details."
            write-Errorlog $logfile $error[0].exception.message $msgData
            Clear-Variable msgData
        }
    if ($TrustedIPs.count -ne 0)
    {
        foreach ($TrustedIP in $TrustedIPs)
        {
            $IP = get-CsTenantTrustedIPAddress | Where-Object {$_.IPAddress -eq $TrustedIP.IPAddress}
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $IP.Identity
            $detail | add-Member -MemberType NoteProperty -Name "IPAddress" -Value $IP.IPAddress
            $detail | add-Member -MemberType NoteProperty -Name "MaskBits" -Value $IP.MaskBits
            $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $IP.Description
            $details += $detail
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "Trusted IP address"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-LISLocation
{
    # Get Emergency Location information Services 
    Write-Host 'Getting Emergency Location Information Services'
    $locations = Get-CsOnlineLisLocation
    if ($locations.count -ne 0)
    {
        $Details = @()
        Foreach ($loc in $locations)
        {
            $detail = New-Object PSObject
            $detail | Add-Member NoteProperty -Name "CompanyName" -Value $loc.CompanyName
            $detail | Add-Member NoteProperty -Name "Civicaddressid" -Value $loc.civicaddressid
            $detail | Add-Member NoteProperty -Name "locationid" -Value $loc.LocationId
            $detail | Add-Member NoteProperty -Name "Description" -Value $loc.Description
            $detail | Add-Member NoteProperty -Name "location" -Value $loc.location
            $detail | Add-Member NoteProperty -Name "HouseNumber" -Value $loc.HouseNumber
            $detail | Add-Member NoteProperty -Name "HouseNumberSuffix" -Value $loc.HouseNumberSuffix
            $detail | Add-Member NoteProperty -Name "PreDirectional" -Value $loc.PreDirectional
            $detail | Add-Member NoteProperty -Name "StreetName" -Value $loc.StreetName
            $detail | Add-Member NoteProperty -Name "PostDirectional" -Value $loc.PostDirectional
            $detail | Add-Member NoteProperty -Name "StreetSuffix" -Value $loc.StreetSuffix
            $detail | Add-Member NoteProperty -Name "City" -Value $loc.City
            $detail | Add-Member NoteProperty -Name "StateOrProvince" -Value $loc.StateOrProvince
            $detail | Add-Member NoteProperty -Name "PostalCode" -Value $loc.PostalCode
            $detail | Add-Member NoteProperty -Name "Country" -Value $loc.CountryOrRegion
            $detail | Add-Member NoteProperty -Name "Latitude" -Value $loc.Latitude
            $detail | Add-Member NoteProperty -Name "Longitude" -Value $loc.Longitude
            $Details += $detail
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "LIS Location"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-LISSubnets
{
    # Get LIS Network information
    Write-Host 'Getting LIS Network Information'
    try {$subnets = Get-CsOnlineLisSubnet -erroraction Stop}
    catch 
        {   
            $msgdata = "Error getting LIS Subnets Details."
            write-Errorlog $logfile $error[0].exception.message $msgData
            Clear-Variable msgData
        }
    if ($subnets.count -ne 0)
        {
        $Details = @()
        Foreach ($subnet in $subnets)
        {
            $detail = New-Object PSObject
            $detail | Add-Member NoteProperty -Name "Subnet" -Value $subnet.Subnet
            $detail | Add-Member NoteProperty -Name "Description" -Value $subnet.Description
            $subloc = Get-CsOnlineLisLocation -LocationId $subnet.LocationId
            $detail | Add-Member NoteProperty -Name "Location" -Value $subloc.location
            $detail | Add-Member NoteProperty -Name "City" -Value $subloc.city
            $Details += $detail
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "LIS Network"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-BSSIDs
{
    #Get LIS Wireless Access Point information
    Write-Host 'Getting LIS WAP Information'
    try {$WAPs = Get-CsOnlineLisWirelessAccessPoint -ErrorAction Stop}
    catch 
    {
        $msgdata = "Error getting LIS WAP Details."
        write-Errorlog $logfile $error[0].exception.message $msgData
        Clear-Variable msgData
    }
    if ($waps.count -ne 0)
    {
        $Details = @()
        Foreach ($WAP in $WAPs)
        {
            $detail = New-Object PSObject
            $detail | Add-Member NoteProperty -Name "BSSID" -Value $WAP.BSSID
            $detail | Add-Member NoteProperty -Name "Description" -Value $WAP.Description
            $WAPloc = Get-CsOnlineLisLocation -LocationId $WAP.LocationId
            $detail | Add-Member NoteProperty -Name "Location" -Value $WAPloc.location
            $detail | Add-Member NoteProperty -Name "City" -Value $WAPloc.city
            $Details += $detail
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "LIS WAP"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-LISSwitch
{
    #Get LIS Switch information
    Write-Host 'Getting LIS SWitch information'
    $Switches = Get-CsOnlineLisSwitch -ErrorAction Stop
    $Details = @()
    if ($Switches.count -ne 0)
    {
        
        Foreach ($Switch in $Switches)
        {
            $detail = New-Object PSObject
            $detail | Add-Member NoteProperty -Name "ChassisID" -Value $Switch.ChassisID
            $detail | Add-Member NoteProperty -Name "Description" -Value $Switch.Description
            $Switchloc = Get-CsOnlineLisLocation -LocationId $Switch.LocationId
            $detail | Add-Member NoteProperty -Name "Location" -Value $Switchloc.location
            $detail | Add-Member NoteProperty -Name "City" -Value $Switchloc.city
            $Details += $detail
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "LIS Switch"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-LISPort
{
     #Get LIS Port information
     Write-Host 'Getting LIS Port Information'
     try {$Ports = Get-CsOnlineLisPort -ErrorAction stop}
     catch 
     {
         $msgdata = "Error getting LIS Port Details."
         write-Errorlog $logfile $error[0].exception.message $msgData
         Clear-Variable msgData
     }
     if ($ports.count -gt 0)
     {
         $Details = @()
         Foreach ($port in $ports)
             {
             $detail = New-Object PSObject
             $detail | Add-Member NoteProperty -Name "ChassisID" -Value $port.ChassisID
             $detail | Add-Member NoteProperty -Name "PortID" -Value $port.PortID
             $detail | Add-Member NoteProperty -Name "Description" -Value $port.Description
             $portloc = Get-CsOnlineLisLocation -LocationId $port.LocationId
             $detail | Add-Member NoteProperty -Name "Location" -Value $portloc.location
             $detail | Add-Member NoteProperty -Name "City" -Value $portloc.city
             $Details += $detail
             }
     }
     else {$details = "No data to display"}
     $tabname = "LIS Port"
     $tabcolor = "Red"
     Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Get-TeamsEnvironment
{
    param ($filelocation, $logname)
    $IncEmployees = Read-host "Include Enterprise Voice Users (y/n)"
    Clear-Host
    Write-Host "Running"
    Write-TenantInfo
    Write-PSTNGateways
    write-PSTNUsages
    Write-VoiceRoutes
    Write-VoiceRoutingPolicies
    Write-TeamsMeetingsSettings
    If ($IncEmployees -eq "y" -or $IncEmployees -eq "Y")
        {Write-EVUsers}
    Else 
        {Write-Host "Skipping Enterprise Voice Users" }
    Write-AutoAttendants 
    Write-CallQueues
    Write-ResourceAccounts
    Write-CallerIDPolicy
    Write-CallingPolicies
    Write-AudioConferencingPolicy
    Write-EmergencyCallingPolicy
    Write-EmergencyCallRouting
    Write-NetworkRegion
    Write-NetworkSiteDetails
    Write-NetworkSubnetDetails
    Write-TrustedIPs
    Write-LISLocation
    Write-LISSubnets
    Write-BSSIDs
    Write-LISSwitch
    Write-LISPort  
    Write-Host "File stored in:" $filelocation
}

Clear-Host
Write-Host "This is will create an Excel Spreadsheet."
$dirlocation = Read-Host "Enter location to store report (i.e. c:\scriptout)"
$directory = $dirlocation+"\TeamsEnvironmentReports"
try { Resolve-Path -Path $directory -ErrorAction Stop }
catch 
    {
        Try {new-item -path $directory -itemtype "Directory" -ErrorAction Stop}
        Catch 
        {
            $logfile, $errordata, $msgData
            $date = get-date -Format "MM/dd/yyyy HH:mm"
            $errordetail = $date + ", Error creating directory. ," + $directory+ ","+ $error[0].exception.message 
            Write-Host $errordetail
        }
    }

Import-Module ImportExcel

# Determine if ImportExcel module is loaded
$XLmodule = Get-Module -Name importexcel
if ($XLmodule )
    {
        If ( $connected=get-cstenant -ErrorAction SilentlyContinue)
            {
                write-host "Current Tenant:" $connected.displayname
                $filedate=Get-Date -Format "MM-dd-yyyy.HH.mm.ss"
                $tenant = $connected.displayname.Replace(" ","-")
                $filelocation = $directory+"\"+$tenant+"-TeamsEnv-"+$filedate+".xlsx"
                $logfile = $directory+"\"+$tenant+"-TeamsEnv-ErrorLog-"+$filedate+".csv"
                Get-TeamsEnvironment $filelocation $logfile
            }
        Else {Write-Host "Teams module isn't loaded.  Please load Teams Module (connect-microsoftteams)"  }
    }
Else {Write-Host "ImportExcel module is not loaded"}