<#
        SRA-T2TDiscovery_1.1.3.PS1
        v1.1.1
        01/12/2024
        simonan@softcat.com        
        
        .SYNOPSIS


        .DESCRIPTION


        .PARAMETER Name
        
        .PARAMETER Extension
        
        .INPUTS
        None. 

        .OUTPUTS
        
        .EXAMPLE


        .LINK
        https://www.softcat.com

    #>

#region Disclaimer

Write-Host "Disclaimer:
This PowerShell script is provided as is without any warranty of any kind, either express or implied, including but not 
limited to the implied warranties of merchantability and fitness for a particular purpose.
The entire risk as to the quality and performance of the script is with you. Should the script prove defective, you assume 
the cost of all necessary servicing, repair, or correction.
In no event shall the author or contributors be liable for any damages whatsoever (including, without limitation, damages 
for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of
the use of or inability to use this script, even if the author has been advised of the possibility of such damages.
Use this script at your own risk."

$continue = Read-Host "Do you want to continue? (Y/N)"
if ($continue -ne "Y") {
    Write-Host "Script execution has been cancelled." -ForegroundColor Red
    Break
}

#endregion Disclaimer

#region Variables and Template Check

$ScriptVersion = "1.1.3"
$TemplateVersion = "1.1.3"
$Template = ""
$Template = "Template_" + $TemplateVersion + ".xlsx"

# Check if the Template exists
if (-not (Test-Path -Path $Template)) {
    Write-Host "The file" $Template "does not exist. Ensure the template is in the working directory and restert the script" -ForegroundColor Red
    Break
}
Write-Host "The file" $Template "exists. Continuing with the script..." -ForegroundColor Green

$Consultant = Read-Host "Please Enter Your Name and press Enter"

#endregion Variables and Template Check

#region Disconnect-GraphSession
Function Disconnect-GraphSession {

    # Disconnect from the Microsoft Graph API 
    Disconnect-MgGraph
    Disconnect-Graph
        
    # Remove the cached Graph API token
    Remove-Item "$env:USERPROFILE\.graph" -Recurse -Force
}
#endregion Disconnect-GraphSession

#region Process Bars
$p = 1
# Update this to the number of process so the count is correct
$TP = 42
# Sets the start time for the elapsed time counter
$StartTime = $(get-date)
#endregion Process Bars

#region Manage Template and Paths
$Org = Get-MgOrganization
$OrgDisplayName = $Org.DisplayName
$OrgID = $Org.Id
$OrgDisplayName = $OrgDisplayName -replace '[\W]', '_'
$Date = Get-Date -Format 'yyyyMMdd_HHmmss'
$TodaysDate = Get-Date
if (-not (Test-Path -Path .\$OrgDisplayName)) {
    Write-Host "The Folder" $Template "does not exist. Creating folder" -ForegroundColor Yellow
    New-Item -Name $OrgDisplayName -ItemType "directory"
}
Else {
        Write-Host "Folder"$OrgDisplayName "exists. Continuing....." -ForegroundColor Green
        }
$Output = $OrgDisplayName + "-" + $Date + ".xlsx"
Write-Host "Copying $Template to $Output"
Copy-Item $Template .\$orgdisplayname\$Output
CD $OrgDisplayName
if (-not (Test-Path -Path .\CSVFiles)) {
    Write-Host "The Folder" $Template "does not exist. Creating folder" -ForegroundColor Yellow
    New-Item -Name "CSVFiles" -ItemType "directory"
}
Else {
        Write-Host "Folder CSVFiles exists. Continuing....." -ForegroundColor Green
        }


$Transcript = $Org.Displayname + "-" + $Date + "-Transcript.Log"
Start-Transcript -Path $Transcript
#endregion Manage Template and Paths

#region Script

#region Populate Teams Variable
$Process = "Populate Teams variable - No Processing"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$AllTeams = @()
$AllTeams = Get-Team | Sort-Object Displayname
Write-Progress  -ID 1 -Activity "Populate Teams variable - No Processing" -Completed
$P++
#endregion Populate Teams Variable

#region Populate SharePoint Variable
$Process = "Populate SharePoint variable - No Processing"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$SPOSites = @()
$SPOSites = Get-SPOSite -IncludePersonalSite $False -Limit All | Where-Object { $_.URL -notlike "*my.sharepoint.com/" -and $_.URL -NotLike "*sharepoint.com/search" } | Sort-Object Title 
Write-Progress  -ID 1 -Activity "Populate SharePoint variable - No Processing" -Completed
$P++
#endregion Populate SharePoint Variable

#region Populate Unified Group Variable
$Process = "Populate Unified Group variable - No Processing"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$UnifiedGroups = @()
$UnifiedGroups = Get-UnifiedGroup -ResultSize Unlimited | Sort-Object DisplayName
Write-Progress  -ID 1 -Activity "Populate UnifiedGroup variable - No Processing" -Completed
$P++
#endregion Populate Unified Group Variable

#region Populate MGUsers Variable
$Process = "All Accounts"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$MGUsers = @()
$Properties = @()
$Properties = @('ID', 'UserPrincipalName', 'DisplayName', 'Mail', 'LicenseAssignmentStates', 'UsageLocation', 'UserType', 'AccountEnabled', 'OnPremisesSyncEnabled', 'SignInActivity')
$MGUsers = Get-MgUser -All -Property $Properties | Sort-Object displayname
Write-Progress  -ID 1 -Activity "Populate All Accounts - No Processing" -Completed
$P++
#endregion Populate MGUsers Variable

#region Populate AllMailboxes Variable
$Process = "All Mailboxes"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$AllMailboxes = @()
$AllMailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.PrimarySMTPAddress -notLike "DiscoverySearchMailbox*" } | Sort-Object UserPrincipalname
Write-Progress  -ID 1 -Activity "Populate All Mailboxes - No Processing" -Completed
$P++
#endregion Populate AllMailboxes Variable

#region Populate MailUser Variable
$Process = "Populate MailUsers"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$AllMailusers = @()
$AllMailusers = Get-mailuser -ResultSize Unlimited
Write-Progress  -ID 1 -Activity "Populate MailUsers - No Processing" -Completed
$P++
#endregion Populate MailUser Variable

#region Populate MGGroups Variable
$Process = "Populate Groups"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$AllMGGroups = @()
$AllMGGroups = Get-MGGroup -All
Write-Progress  -ID 1 -Activity "Populate Groups - No Processing" -Completed
$P++
#endregion Populate MailUser Variable

#region Registered Applications
$Process = "Registered Applications"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)

$MGRegisteredApplications = Get-MgApplication
$MGRegisteredApplicationsArray = @()
$i = 1
Foreach ($MGRegisteredApplication in $MGRegisteredApplications) {
    If ($MGRegisteredApplications.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Registered Applications" -Status "Environment $i of $($MGRegisteredApplications.Count)" -PercentComplete (($i / $MGRegisteredApplications.Count) * 100) 
    }
    $MGRegisteredApplicationsArray = $MGRegisteredApplicationsArray + [PSCustomObject]@{
        DisplayName     = $MGRegisteredApplication.DisplayName ;
        AppId           = $MGRegisteredApplication.AppId ;
        PublisherDomain = $MGRegisteredApplication.PublisherDomain ;
        CreatedDateTime = $MGRegisteredApplication.CreatedDateTime ;
        ORGID           = $OrgID
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If ($MGRegisteredApplicationsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MGRegisteredApplicationsArray | Export-Excel -Path $Output -AutoSize -TableName Registered_Applications -WorksheetName Registered_Applications  
}
Write-Progress  -ID 1 -Activity "Processing Registered Applications" -Completed
$P++
#endregion Registered Applications

#region Power Environments
$Process = "Power Environments"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$PowerEnvironments = Get-AdminPowerAppEnvironment | Sort DisplayName
$PowerEnvironmentsArray = @()
$i = 1
Foreach ($PowerEnvironment in $PowerEnvironments) {
    If ($PowerEnvironments.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Power Apps" -Status "Environment $i of $($PowerEnvironments.Count)" -PercentComplete (($i / $PowerEnvironments.Count) * 100) 
    }
    $PowerEnvironmentsArray = $PowerEnvironmentsArray + [PSCustomObject]@{
        EnvironmentName = $PowerEnvironment.EnvironmentName ;
        DisplayName     = $PowerEnvironment.DisplayName ;
        IsDefault       = $PowerEnvironment.IsDefault ;
        Location        = $PowerEnvironment.Location ;
        Created         = $PowerEnvironment.CreatedTime ;
        CreatedBy       = $PowerEnvironment.CreatedBy.DisplayName ;
        CreationType    = $PowerEnvironment.CreationType ;
        EnvironmentType = $PowerEnvironment.EnvironmentType ;
        Description     = $PowerEnvironment.Description ;
        ORGID           = $OrgID
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If ($PowerEnvironmentsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $PowerEnvironmentsArray | Export-Excel -Path $Output -AutoSize -TableName Power_Environments -WorksheetName Power_Environments  
}
Write-Progress  -ID 1 -Activity "Processing Power Environments" -Completed
$P++
#endregion Power Environments

#region Power Apps
$Process = "Power Apps"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$PowerApps = Get-AdminPowerApp | Sort DisplayName
$PowerAppsArray = @()
$i = 1
Foreach ($PowerApp in $PowerApps) {
    If ($PowerApps.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Power Apps" -Status "Power App $i of $($PowerApps.Count)" -PercentComplete (($i / $PowerApps.Count) * 100) 
    }
    $PowerAppsArray = $PowerAppsArray + [PSCustomObject]@{
        DisplayName     = $PowerApp.DisplayName ;
        AppType         = $PowerApp.Internal.appType ;
        Created         = $PowerApp.CreatedTime ;
        EnvironmentName = $PowerApp.EnvironmentName ;
        Owner           = $Powerapp.owner.displayName ;
        OwnerEmail      = $Powerapp.owner.email ;
        ORGID           = $OrgID
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If ($PowerAppsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $PowerAppsArray | Export-Excel -Path $Output -AutoSize -TableName Power_Apps -WorksheetName Power_Apps  
}
Write-Progress  -ID 1 -Activity "Processing Power Apps" -Completed
$P++
#endregion Power Apps

#region Power Flows
$Process = "Power Flows"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$PowerFlows = Get-AdminFlow | Sort DisplayName
$PowerFlowsArray = @()
$i = 1
Foreach ($PowerFlow in $PowerFlows) {
    If ($PowerFlows.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Power Flows" -Status "Flow $i of $($PowerFlows.Count)" -PercentComplete (($i / $PowerFlows.Count) * 100) 
    }
    $PowerFlowsArray = $PowerFlowsArray + [PSCustomObject]@{
        DisplayName     = $PowerFlow.DisplayName ;
        Enabled         = $PowerFlow.Enabled ;
        UserType        = $PowerFlow.UserType ;
        CreatedTime     = $PowerFlow.CreatedTime ;
        CreatedBy       = $PowerFlow.CreatedBy.UserID ;
        EnvironmentName = $PowerFlow.EnvironmentName ;
        ORGID           = $OrgID
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If ($PowerFlowsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $PowerFlowsArray | Export-Excel -Path $Output -AutoSize -TableName Power_Flows -WorksheetName Power_Flows  
}
Write-Progress  -ID 1 -Activity "Processing Power Flows" -Completed
$P++
#endregion Power Flows

#region Power BI Reports
$Process = "Power BI Reports"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$PowerBIReports = Get-PowerBIReport -Scope Organization | Sort DisplayName
$PowerBIReportsArray = @()
$i = 1
Foreach ($PowerBIReport in $PowerBIReports) {
    If ($PowerBIReports.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Power BI Reports" -Status "Power BI Report $i of $($PowerBIReports.Count)" -PercentComplete (($i / $PowerBIReports.Count) * 100) 
    }
    $PowerBIReportsArray = $PowerBIReportsArray + [PSCustomObject]@{
        Name  = $PowerBIReport.Name ;
        ID    = $PowerBIReport.ID ;
        ORGID = $OrgID
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If ($PowerBIReportsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $PowerBIReportsArray | Export-Excel -Path $Output -AutoSize -TableName Power_BI_Reports -WorksheetName Power_BI_Reports  
}
Write-Progress  -ID 1 -Activity "Processing Power BI Reports" -Completed
$P++
#endregion Power BI Reports

#region Bitlocker Keys
$Process = "Bitlocker Keys"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$BitlockerKeys = Get-MgInformationProtectionBitlockerRecoveryKey
$BitlockerKeysArray = @()
$i = 1
Foreach ($BitlockerKey in $BitlockerKeys) {
    If ($BitlockerKeys.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Bitlocker Keys" -Status "Bitlocker Key $i of $($BitlockerKeys.Count)" -PercentComplete (($i / $BitlockerKeys.Count) * 100) 
    }
    $BitlockerKeysArray = $BitlockerKeysArray + [PSCustomObject]@{
        CreatedDateTime  = $BitlockerKey.CreatedDateTime ;
        DeviceId    = $BitlockerKey.DeviceId ;
        ID    = $BitlockerKey.ID ;
        VolumeType = $BitlockerKey.VolumeType
        ORGID = $OrgID
    }
$i++
}
Start-Sleep 5
If ($BitlockerKeysArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $BitlockerKeysArray | Export-Excel -Path $Output -AutoSize -TableName Bitlocker_Keys -WorksheetName Bitlocker_Keys  
}
Write-Progress  -ID 1 -Activity "Processing Bitlocker Keys" -Completed
$P++


#endregion Bitlocker Keys

#region Power BI Dashboards
$Process = "Power BI Dashboards"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$PowerBIDashboards = Get-PowerBIDashboard -Scope Organization | Sort DisplayName
$PowerBIDashboardsArray = @()
$i = 1
Foreach ($PowerBIDashboard in $PowerBIDashboards) {
    If ($PowerBIDashboards.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Power BI Dashboards" -Status "Power BI Report $i of $($PowerBIDashboards.Count)" -PercentComplete (($i / $PowerBIDashboards.Count) * 100) 
    }
    $PowerBIDashboardsArray = $PowerBIDashboardsArray + [PSCustomObject]@{
        Name  = $PowerBIDashboard.Name ;
        ID    = $PowerBIDashboard.ID ;
        ORGID = $OrgID
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If ($PowerBIDashboardsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $PowerBIDashboardsArray | Export-Excel -Path $Output -AutoSize -TableName Power_BI_Dashboards -WorksheetName Power_BI_Dashboards  
}
Write-Progress  -ID 1 -Activity "Processing Power BI Dashboards" -Completed
$P++
#endregion Power BI Dashboards

#region All MGAccounts
$MGUsersArray = @()
$i = 1
Foreach ($MGUser in $MGUsers) {
    If ($MGUsers.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing All MG User Accounts" -Status "User Account $i of $($MGUsers.Count)" -PercentComplete (($i / $MGUsers.Count) * 100) 
    }
   $LastLoginDate = ""
    $DaysDifference = ""
    $LastLoginDate = $MGUser.SignInActivity.LastSignInDateTime
    If ($LastLoginDate -ne $Null) {
        $DaysDifference = ($TodaysDate - $LastLoginDate).Days
    }
    $LicenseSKUs = Get-MgUserLicenseDetail -UserId $MGUser.ID 
    
    $MGUsersArray = $MGUsersArray + [PSCustomObject]@{
        ID                 = $MGUser.ID;
        DisplayName        = $MGUser.DisplayName ;
        UserPrincipalName  = $MGUser.UserPrincipalName ;
        Mail               = $MGUser.Mail ;
        UserType           = $MGUser.UserType ;
        UsageLocation      = $MGUser.UsageLocation ;
        AccountEnabled     = $MGUser.AccountEnabled ;
        LicenseSKUs        = $LicenseSKUs.SkuPartNumber -join ";" ;
        IsDirSynced        = $MGUser.OnPremisesSyncEnabled ;
        LastLoginDate      = $LastLoginDate ;
        DaysSinceLastLogin = $DaysDifference ;
        ORGID              = $OrgID
    }
    $i++
}
Start-Sleep 5
If ($MGUsersArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"    
    $MGUsersArray | Export-Excel -Path $Output -AutoSize -TableName Accounts_AllMGUsers -WorksheetName Accounts_AllMGUsers 
    $MGUsersArray | Export-csv  .\csvfiles\Accounts_AllMGUsers.csv -notypeinformation
    $MGUsers = @()
}
Write-Progress  -ID 1 -Activity "Processing All MG User Accounts" -Completed
$P++
#endregion All MGAccounts

#region Licensed MGAccounts
$Process = "Licensed MGAccounts"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$MGLicensedAccounts = $MGUsersArray | Where { $_.LicenseSKUs -ne "" }
$MGLicensedAccountsArray = @()
$i = 1
Foreach ($MGLicensedAccount in $MGLicensedAccounts) {
    If ($MGLicensedAccounts.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Licensed Accounts" -Status "Licensed Account $i of $($MGLicensedAccounts.Count)" -PercentComplete (($i / $MGLicensedAccounts.Count) * 100) 
    }
    $MGLicensedAccountsArray = $MGLicensedAccountsArray + [PSCustomObject]@{
        ID                 = $MGLicensedAccount.ID;
        DisplayName        = $MGLicensedAccount.DisplayName ;
        UserPrincipalName  = $MGLicensedAccount.UserPrincipalName ;
        Mail               = $MGLicensedAccount.Mail ;
        UserType           = $MGLicensedAccount.UserType
        UsageLocation      = $MGLicensedAccount.UsageLocation ;
        AccountEnabled     = $MGLicensedAccount.AccountEnabled ;
        LicenseSKUs        = $MGLicensedAccount.LicenseSKUs -join ";" ;
        IsDirSynced        = $MGLicensedAccount.IsDirSynced ;
        LastLoginDate      = $LastLoginDate ;
        DaysSinceLastLogin = $DaysDifference ;
        ORGID              = $OrgID
    }
    $1++
}
#Start-Sleep 5
If ($MGLicensedAccountsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MGLicensedAccountsArray | Export-Excel -Path $Output -AutoSize -TableName Accounts_Licensed -WorksheetName Accounts_Licensed  
    $MGLicensedAccountsArray | Export-CSV .\csvfiles\Accounts_Licensed.CSV -notypeinformation 
}
Write-Progress  -ID 1 -Activity "Processing Licensed Accounts" -Completed
$P++

#endregion Licensed MGAccounts

#region UnLicensed MGAccounts
$Process = "UnLicensed Accounts"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$MGUnLicensedAccounts = $MGUsersArray | Where { $_.LicenseSKUs -eq "" }
$MGUnLicensedAccountsArray = @()
$i = 1
Foreach ($MGUnLicensedAccount in $MGUnLicensedAccounts) {
    If ($MGLicensedAccounts.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing UnLicensed Accounts" -Status "UnLicensed Account $i of $($MGUnLicensedAccounts.Count)" -PercentComplete (($i / $MGUnLicensedAccounts.Count) * 100) 
    }
    $MGUnLicensedAccountsArray = $MGUnLicensedAccountsArray + [PSCustomObject]@{
        ID                 = $MGUnLicensedAccount.ID;
        DisplayName        = $MGUnLicensedAccount.DisplayName ;
        UserPrincipalName  = $MGUnLicensedAccount.UserPrincipalName ;
        Mail               = $MGUnLicensedAccount.Mail ;
        UserType           = $MGUnLicensedAccount.UserType
        UsageLocation      = $MGUnLicensedAccount.UsageLocation ;
        AccountEnabled     = $MGUnLicensedAccount.AccountEnabled ;
        LicenseSKUs        = $MGUnLicensedAccount.LicenseSKUs -join ";" ;
        IsDirSynced        = $MGUnLicensedAccount.IsDirSynced;
        LastLoginDate      = $LastLoginDate ;
        DaysSinceLastLogin = $DaysDifference ;
        ORGID              = $OrgID
    }
    $1++
}
#Start-Sleep 5
If ($MGUnLicensedAccountsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MGUnLicensedAccountsArray | Export-Excel -Path $Output -AutoSize -TableName Accounts_UnLicensed -WorksheetName Accounts_UnLicensed  
    $MGUnLicensedAccountsArray | Export-CSV .\csvfiles\Accounts_UnLicensed.CSV -notypeinformation 
}
Write-Progress  -ID 1 -Activity "Processing UnLicensed Accounts" -Completed
$P++
#endregion UnLicensed MGAccounts

#region Contacts
$Process = "Contacts"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$MGContactsArray = @()
$MGContacts = Get-MGContact -All | Sort-Object Displayname
$i = 1
Foreach ($MGContact in $MGContacts) { 
    If ($MGContacts.Count -gt "1") { 
        Write-Progress  -ID 1 -Activity "Processing Contacts" -Status "MGContact $i of $($MGContacts.Count)" -PercentComplete (($i / $MGContacts.Count) * 100)  
    }
    $MGContactsArray = $MGContactsArray + [PSCustomObject]@{
        DisplayName = $MGContact.DisplayName ; 
        #RecipientTypeDetails = $MGContact.RecipientTypeDetails ;
        Company     = $MGContact.Company ; 
        FirstName   = $MGContact.GivenName ; 
        LastName    = $MGContact.SurName ; 
        Email       = $MGContact.Mail ; 
        IsDirSynced = $MGContact.OnPremisesSyncEnabled ;
        ORGID       = $OrgID 
    }
    $i++
}
Start-Sleep 5
If ($MGContactsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"    
    $MGContactsArray | Export-Excel -Path $Output -AutoSize -TableName Accounts_Contacts -WorksheetName Accounts_Contacts
    $MGContactsArray | Export-CSV .\csvfiles\Accounts_Contacts.CSV -NoTypeInformation
}
Write-Progress  -ID 1 -Activity "Processing Contacts" -Completed
$P++
#endregion Contacts

#region Account SKUs
$Process = "License SKUs"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)

$MGAccountSkus = Get-MgSubscribedSku | Sort-Object SkuPartNumber 
$MGSKUArray = @()
$i = 1
Foreach ($MGAccountSku in $MGAccountSkus) {
    If ($MGAccountSkus.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing License SKUs" -Status "SKU $i of $($MGAccountSkus.Count)" -PercentComplete (($i / $MGAccountSkus.Count) * 100) 
    }
    $SKU = (get-MgSubscribedSku -SubscribedSkuId $MGAccountSku.id).SkuPartNumber
    $LicenseCount = (get-MgSubscribedSku -SubscribedSkuId $MGAccountSku.id -Property PrepaidUnits | select-object -expandproperty prepaidunits).enabled
    $ConsumedLicenses = (get-MgSubscribedSku -SubscribedSkuId $MGAccountSku.id).ConsumedUnits
    $MGSKUArray = $MGSKUArray + [PSCustomObject]@{
        SKU       = $SKU;
        Purchased = $LicenseCount ;
        Consumed  = $ConsumedLicenses ; 
        Available = $LicenseCount - $ConsumedLicenses ;
        ORGID     = $OrgID
    }
    #    Start-Sleep 1
    $i++
}
If ($MGSKUArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MGSKUArray | Export-Excel -Path $Output -AutoSize -TableName Accounts_SKU -WorksheetName Accounts_SKU  
    $MGSKUArray | Export-CSV .\csvfiles\Accounts_SKU.CSV -NoTypeInformation 
}
Write-Progress  -ID 1 -Activity "Processing License SKUs" -Completed
$P++
#endregion Account SKUs

#region OneDrives
$Process = "OneDrives"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets all OneDrive data
$OneDrives = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'" | Sort-Object Title
$OneDriveArray = @()
$i = 1
Foreach ($OneDrive in $OneDrives) {
    If ($OneDrives.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing OneDrives" -Status "OneDrive $i of $($OneDrives.Count)" -PercentComplete (($i / $OneDrives.Count) * 100)  
    }
    $StorageUsageCurrentGB = ""
    $StorageQuotaGB = ""
    $StorageUsageCurrentGB = $OneDrive.StorageUsageCurrent / 1024
    $StorageQuotaGB = $OneDrive.StorageQuota / 1024

    $OneDriveArray = $OneDriveArray + [PSCustomObject]@{
        Status                       = $OneDrive.Status ;
        Title                        = $OneDrive.Title ; 
        Owner                        = $OneDrive.Owner ;
        LastContentModifiedDate      = $OneDrive.LastContentModifiedDate ; 
        StorageUsageCurrentGB        = $StorageUsageCurrentGB ;
        StorageQuotaGB               = $StorageQuotaGB ;  
        Url                          = $OneDrive.Url ;
        SharingCapability            = $OneDrive.SharingCapability ; 
        SiteDefinedSharingCapability = $OneDrive.SiteDefinedSharingCapability ; 
        ConditionalAccessPolicy      = $OneDrive.ConditionalAccessPolicy ;
        ORGID                        = $OrgID 
    }
    $i++
}
Start-Sleep 5
If ($OneDriveArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $OneDriveArray | Export-Excel -Path $Output -AutoSize -TableName OneDrives -WorksheetName OneDrives
    $OneDriveArray | Export-CSV .\csvfiles\OneDrives.csv -NoTypeInformation           
}
Start-Sleep 5
$LargestOneDrives = $OneDriveArray | Sort-Object StorageUsageCurrentMB -Descending | Select-Object -First 10
If ($LargestOneDrives -ne $Null) {
    $LargestOneDrives | Export-Excel -Path $Output -AutoSize -TableName OneDrives_TopTen -WorksheetName OneDrives_TopTen
}
$OneDrives = @()
Write-Progress  -ID 1 -Activity "Processing OneDrives" -Completed
$P++
#endregion OneDrives

#region All MGGroups
$MGGroupsArray = @()
$i = 1
Foreach ($MGGroup in $AllMGGroups) {
    If ($MGGroups.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing All MG Groups" -Status "Group $i of $($AllMGGroups.Count)" -PercentComplete (($i / $AllMGGroups.Count) * 100) 
    }
    $Type = ""
    $Grouptype = ""
    $Grouptype = $MGGroup.GroupTypes -join ";"
    If ($Grouptype -Like "*Unified*") {
        $Type = "Unified"
    }
    If ($Grouptype -NotLike "*Unified*" -and $MGGroup.MailEnabled -eq $False -and $MGGroup.SecurityEnabled -eq $True) {
        $Type = "Security"
    }
    If ($Grouptype -eq "" -and $MGGroup.MailEnabled -eq $False -and $MGGroup.SecurityEnabled -eq $True) {
        $Type = "Security"
    }
    If ($Grouptype -eq "" -and $MGGroup.MailEnabled -eq $True -and $MGGroup.SecurityEnabled -eq $True) {
        $Type = "Mail Enabled Security"
    }
    If ($Grouptype -eq "" -and $MGGroup.MailEnabled -eq $True -and $MGGroup.SecurityEnabled -eq $False) {
        $Type = "Distribution"
    }
    $MGGroupsArray = $MGGroupsArray + [PSCustomObject]@{
        ID              = $MGGroup.ID;
        DisplayName     = $MGGroup.DisplayName ;
        Type            = $Type
        MailEnabled     = $MGGroup.MailEnabled ;
        SecurityEnabled = $MGGroup.SecurityEnabled ;
        GroupTypes      = $MGGroup.GroupTypes -join ";" ;
        IsDirSynced     = $MGGroup.OnPremisesSyncEnabled ;
        ORGID           = $OrgID
    }
    $i++
}
Start-Sleep 5
If ($MGGroupsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"    
    $MGGroupsArray | Export-Excel -Path $Output -AutoSize -TableName Groups_MG -WorksheetName Groups_MG 
    $MGGroupsArray | Export-csv  .\csvfiles\Groups_MG.csv -notypeinformation
    $MGUsers = @()
}
Write-Progress  -ID 1 -Activity "Processing All MG Groups" -Completed
$P++
#endregion All MGGroups

#region Devices
$Process = "Devices"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$TodaysDate = Get-Date
# Gets Device data
$MGDevices = Get-MGDevice -All  | Sort-Object Displayname
$MGDeviceArray = @()
$i = 1
Foreach ($MGDevice in $MGDevices) {
    If ($MGdevices.count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Devices" -Status "Device $i of $($MGDevices.Count)" -PercentComplete (($i / $MGDevices.Count) * 100)  
    }
    If ($MGDevice.TrustType -eq "AzureAD") {
        $TrustType = "Microsoft Entra Joined"
    }
    If ($MGDevice.TrustType -eq "ServerAD") {
        $TrustType = "Microsoft Entra Hybrid Joined"
    }
    If ($MGDevice.TrustType -eq "Workplace") {
        $TrustType = "Microsoft Entra Registered"
    }
    $DaysDifference = ""
    If($MGDevice.ApproximateLastSignInDateTime -ne $Null) {
        $DaysDifference = ($TodaysDate - $MGDevice.ApproximateLastSignInDateTime).Days
    }
    $BitlockerInfo = ""
    $Bitlockered = ""
    $BitlockerInfo = $BitlockerKeys | Where{$_.DeviceID -eq $MGDevice.DeviceID}
    If($BitlockerInfo -ne $Null) {
        $Bitlockered = "TRUE"
    }
    $MGDeviceArray = $MGDeviceArray + [PSCustomObject]@{
        DisplayName                   = $MGDevice.DisplayName ;         
        DeviceOsType                  = $MGDevice.OperatingSystem ;
        DeviceOsVersion               = $MGDevice.OperatingSystemVersion ; 
        DeviceTrustType               = $TrustType ; 
        EnrollmentType                = $MGDevice.EnrollmentType ;
        Bitlockered                   = $Bitlockered
        ManagementType                = $MGDevice.ManagementType ;
        Manufacturer                  = $MGDevice.Manufacturer ;
        Model                         = $MGDevice.Model ;
        ApproximateLastLogonTimestamp = $MGDevice.ApproximateLastSignInDateTime ; 
        DaysSinceLastLogin            = $DaysDifference ;
        ORGID                         = $OrgID
    }
    $i++
}
Start-Sleep 5
If ($MGDeviceArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MGDeviceArray | Export-Excel -Path $Output -AutoSize -TableName Devices_All -WorksheetName Devices_All
}
Start-Sleep 5
$MGDevicesUniqueArray = $MGDeviceArray | Sort-Object -Property DisplayName -Unique
If ($MGDevicesUniqueArray -ne $Null) {
    Write-Host "Writing Unique $Process data to $Output"
    $MGDevicesUniqueArray | Export-Excel -Path $Output -AutoSize -TableName Devices_Unique -WorksheetName Devices_Unique
    $MGDevicesUniqueArray | Export-CSV .\csvfiles\Devices_Unique.csv -NoTypeInformation
}
Write-Progress  -ID 1 -Activity "Processing Devices" -Completed
$P++
#endregion Devices

#region Domains
$Process = "Domains"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$AllMGDomains = Get-MGDomain | Sort-Object Name
$AllMGDomainsArray = @()
$i = 1
Foreach ($MGDomain in $AllMGDomains) {
    If ($AllMGDomains.Count -gt "1") {
        Write-Progress -ID 1 -Activity "Processing MG Domains" -Status "Domain $i of $($AllMGDomains.Count)" -PercentComplete (($i / $AllMGDomains.Count) * 100)  
    }
    $AllMGDomainsArray = $AllMGDomainsArray + [PSCustomObject]@{
        Id                 = $MGDomain.Id ;
        AuthenticationType = $MGDomain.AuthenticationType ;
        IsVerified         = $MGDomain.IsVerified ;
        IsDefault          = $MGDomain.IsDefault ;
        IsInitial          = $MGDomain.IsInitial ;
        SupportedServices  = $MGDomain.SupportedServices -join ";" ;
        ORGID              = $OrgID 
    }
    #Start-Sleep 1
    $i++
}
Start-Sleep 5
If ($AllMGDomainsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $AllMGDomainsArray | Export-Excel -Path $Output -AutoSize -TableName Domains_All -WorksheetName Domains_All  
}
Write-Progress -ID 1 -Activity "Processing MG Domains" -Completed
$P++
#endregion Domains

#region MX Records
$Process = "MX Records"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Accepted Domains and resolves MX records
$AcceptedDomains = Get-AcceptedDomain | Sort-Object ascending
$AllMXRecordsArray = @()
$i = 1
$ErrorActionPreference = "SilentlyContinue"
Foreach ($AcceptedDomain in $AcceptedDomains) {
    If ($AcceptedDomains.Count -gt "1") {
        Write-Progress -ID 1 -Activity "Processing MX Records" -Status "Accepted Domain $i of $($AcceptedDomains.Count)" -PercentComplete (($i / $AcceptedDomains.Count) * 100)  
    }
    $MXRecords = $AcceptedDomain | Resolve-DNSName -Type MX -Server 8.8.8.8 | Where-Object { $_.QueryType -eq "MX" }  | Select-Object Name, NameExchange, Preference, TTL | Sort-Object Preference 
    Foreach ($MXrecord in $MXRecords) {        
        $AllMXRecordsArray = $AllMXRecordsArray + [PSCustomObject]@{
            Name         = $MXRecord.Name ;
            NameExchange = $MXRecord.NameExchange ;
            Preference   = $MXRecord.Preference ;
            TTL          = $MXRecord.TTL ;
            ORGID        = $OrgID
        }
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If ($AllMXREcordsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $AllMXRecordsArray | Export-Excel -Path $Output -AutoSize -TableName Domains_MXRecords -WorksheetName Domains_MXRecords 
} 
Write-Progress -ID 1 -Activity "Processing MX Records" -Completed
$P++
#endregion MX Records

#region SPF Records
$Process = "SPF Records"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
#Remove after testing
#$AcceptedDomains = Get-AcceptedDomain
$AllSPFRecordsArray = @()
$i = 1
Foreach ($AcceptedDomain in $AcceptedDomains) {
    If ($AcceptedDomains.Count -gt "1") {
        Write-Progress -ID 1 -Activity "Processing SPF Records" -Status "Accepted Domain $i of $($AcceptedDomains.Count)" -PercentComplete (($i / $AcceptedDomains.Count) * 100)  
    }
    $SPFRecords = $AcceptedDomain | resolve-dnsname -Type TXT -Server 8.8.8.8
    Foreach ($SPFRecord in $SPFRecords) {  
        If ($SPFRecord.Strings -like "V=SPF*") {     
            $AllSPFRecordsArray = $AllSPFRecordsArray + [PSCustomObject]@{
                Name   = $SPFRecord.Name ;
                String = $SPFRecord.Strings -join "," ;
                TTL    = $SPFRecord.TTL ;
                ORGID  = $OrgID
            }
        }
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If ($AllSPFRecordsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $AllSPFRecordsArray | Export-Excel -Path $Output -AutoSize -TableName Domains_SPFRecords -WorksheetName Domains_SPFRecords 
} 
Write-Progress -ID 1 -Activity "Processing SPF Records" -Completed
$P++
#endregion SPF Records

#region DMARC Records
$Process = "DMARC Records"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)

$AllDMARCecordsArray = @()
$i = 1
Foreach ($AcceptedDomain in $AcceptedDomains) {
    If ($AcceptedDomains.Count -gt "1") {
        Write-Progress -ID 1 -Activity "Processing DMARC Records" -Status "Accepted Domain $i of $($AcceptedDomains.Count)" -PercentComplete (($i / $AcceptedDomains.Count) * 100)  
    }
    $Domainname = $AcceptedDomain.DomainName
    $params = @{
        Name        = "_dmarc.$DomainName"
        ErrorAction = "SilentlyContinue"
    }
    $dnsTxt = Resolve-DnsName @params -Type  TXT | Where-Object Type -eq TXT  
    #$dnsTxt | Select-Object @{Name = "DMARC"; Expression = {"$DomainName`:$s"}},@{Name = "Record"; Expression = {$_.Strings}} 
    $AllDMARCecordsArray = $AllDMARCecordsArray + [PSCustomObject]@{
        Domain  = $Domainname  ;
        Name    = $dnsTxt.Name ;
        Section = $dnsTxt.Section ;
        String  = $dnsTxt.Strings -join "," ;
        TTL     = $dnsTxt.TTL ;
        ORGID   = $OrgID
    }
        
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If ($AllDMARCecordsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $AllDMARCecordsArray | Export-Excel -Path $Output -AutoSize -TableName Domains_DMARCRecords -WorksheetName Domains_DMARCRecords 
} 
Write-Progress -ID 1 -Activity "Processing DMARC Records" -Completed
$P++
#endregion DMARC Records

#region DKIM Records
$Process = "DKIM Records"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$DKIMConfig = Get-DkimSigningConfig
$DKIMConfigArray = @()
$i = 1
Foreach ($DKIM in $DKIMConfig) {
    If ($DKIMConfig.Count -gt "1") {
        Write-Progress -ID 1 -Activity "Processing DKIM Records" -Status "Record $i of $($DKIMConfig.Count)" -PercentComplete (($i / $DKIMConfig.Count) * 100)  
    }
    $DKIMConfigArray = $DKIMConfigArray + [PSCustomObject]@{
        Domain         = $DKIM.Domain ;
        Enabled        = $DKIM.Enabled ;
        Selector1CNAME = $DKIM.Selector1CNAME ;
        ORGID          = $OrgID 
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If ($DKIM -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $DKIMConfigArray | Export-Excel -Path $Output -AutoSize -TableName Domains_DKIMRecords -WorksheetName Domains_DKIMRecords 
} 
Write-Progress -ID 1 -Activity "Processing DKIM Records" -Completed
$P++
#endregion DKIM Records

#region Recipient UPN Count
$Process = "Recipient Counts"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$i = 1
$LicensedADRecipientsArray = @()
$UnLicensedADRecipientsArray = @()

Foreach ($Accepteddomain in $AcceptedDomains) {
    If ($AcceptedDomains.Count -gt "1") {
        Write-Progress -ID 1 -Activity "Processing Recipient Counts" -Status "Accepted Domain $i of $($AcceptedDomains.Count)" -PercentComplete (($i / $AcceptedDomains.Count) * 100)  
    }
    $domainsuffix = "*@" + $Accepteddomain.Name
    $LicensedADRecipients = $MGLicensedAccountsArray | where { $_.UserPrincipalname -like $domainsuffix }
    $UnLicensedADRecipients = $MGUnLicensedAccountsArray | where { $_.UserPrincipalname -like $domainsuffix }
    $LicensedADRecipientsArray = $LicensedADRecipientsArray + [PSCustomObject]@{
        Suffix     = $domainsuffix ;
        Licensed   = $LicensedADRecipients.Count ;
        UnLicensed = $UnLicensedADRecipients.Count ;
        Total      = $LicensedADRecipients.Count + $UnLicensedADRecipients.Count ;
        ORGID      = $OrgID
    }
    $i++
}
Start-Sleep 5
If ($LicensedADRecipientsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $LicensedADRecipientsArray | Export-Excel -Path $Output -AutoSize -TableName RecipientUPN_Count -WorksheetName RecipientUPN_Count
}
Write-Progress  -ID 1 -Activity "Processing Recipient Counts" -Completed
$P++

#endregion Recipient UPN Count

#region EOL Inbound Connectors
$Process = "Exchange InBound Connectors"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Inbound Connectors
$InboundConnectors = Get-InboundConnector | Sort-Object Name
$InboundConnectorsArray = @()
$i = 1
Foreach ($InboundConnector in $InboundConnectors) {
    If ($InboundConnectors.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Inbound Connectors" -Status "Connector $i of $($InboundConnectorsArray.Count)" -PercentComplete (($i / $InboundConnectorsArray.Count) * 100) 
    }
    $InboundConnectorsArray = $InboundConnectorsArray + [PSCustomObject]@{
        Name              = $InboundConnector.Name; 
        Enabled           = $InboundConnector.Enabled ;
        ConnectorType     = $InboundConnector.ConnectorType ;
        SenderIPAddresses = $InboundConnector.SenderIPAddresses -Join ","  ;
        SenderDomains     = $InboundConnector.SenderDomains -Join "," ;
        ORGID             = $OrgID
    }
    Start-Sleep 1
    $1++
}
Start-Sleep 5
If ($InboundConnectorsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $InboundConnectorsArray | Export-Excel -Path $Output -AutoSize -TableName EOL_InboundConnectors -WorksheetName EOL_InboundConnectors
}
Write-Progress  -ID 1 -Activity "Processing Inbound Connectors" -Completed
$P++
#endregion EOL Inbound Connectors

#region EOL Outbound Connectors
$Process = "Exchange Outbound Connectors"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Outbound Connectors. Does not get TestModeConnectors
$OutboundConnectors = Get-OutboundConnector | Sort-Object Name
$OutboundConnectorsArray = @()
$i = 1
Foreach ($OutboundConnector in $OutboundConnectors) {
    If ($OutboundConnectors.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Outbound Connectors" -Status "Connector $i of $($OutboundConnectorsArray.Count)" -PercentComplete (($i / $OutboundConnectorsArray.Count) * 100) 
    }
    $OutboundConnectorsArray = $OutboundConnectorsArray + [PSCustomObject]@{
        Name             = $OutboundConnector.Name; 
        Enabled          = $OutboundConnector.Enabled ;
        ConnectorType    = $OutboundConnector.ConnectorType ;
        UseMXRecord      = $OutboundConnector.UseMXRecord ;
        IsValidated      = $OutboundConnector.IsValidated ;
        TlsSettings      = $OutboundConnector.TlsSettings ; 
        SmartHosts       = $OutboundConnector.SmartHosts -Join "," ; 
        RecipientDomains = $OutboundConnector.RecipientDomains -Join "," ;
        ORGID            = $OrgID
    }
    Start-Sleep 1
    $1++
}
Start-Sleep 5
If ($OutboundConnectorsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $OutboundConnectorsArray | Export-Excel -Path $Output -AutoSize -TableName EOL_OutboundConnectors -WorksheetName EOL_OutboundConnectors #  -Append
}
Write-Progress  -ID 1 -Activity "Processing Outbound Connectors" -Completed
$P++
#endregion EOL Outbound Connectors

#region Mail Flow Rules
$Process = "Mail Flow Rules"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Transport Rules
$TransportRules = Get-TransportRule | Sort-Object Priority
$TransportRulesArray = @()
$i = 1
ForEach ($TransportRule in $TransportRules) {
    If ($TransportRules.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Transport Rules" -Status "Transport Rule $i of $($TransportRules.Count)" -PercentComplete (($i / $TransportRules.Count) * 100)  
    }
    $TransportRulesArray = $TransportRulesArray + [PSCustomObject]@{
        Name     = $TransportRule.Name ;
        State    = $TransportRule.State ; 
        Mode     = $TransportRule.Mode ;
        Priority = $TransportRule.Priority ; 
        Comments = $TransportRule.Comments ;
        ORGID    = $OrgID
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If ($TransportRulesArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $TransportRulesArray | Export-Excel -Path $Output -AutoSize -TableName EOL_TransportRules -WorksheetName EOL_TransportRules
}
Write-Progress  -ID 1 -Activity "Processing Transport Rules" -Completed
$P++
#endregion Mail Flow Rules

#region Distribution Groups
$DistributionGroups = Get-DistributionGroup -ResultSize Unlimited | Sort-Object DisplayName
$DGMemberArray = @()
$DistributionGroupsArray = @()
$DistributionGroupsMembersArray = @()
$ExternalDistributionGroupMember = @()
$ExternalDGMemberArray = @()
$i = 1
Foreach ($DistributionGroup in $DistributionGroups) {
    If ($DistributionGroups.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Distribution Groups" -Status "Distribution Group $i of $($DistributionGroups.Count)" -PercentComplete (($i / $DistributionGroups.Count) * 100) 
    }
    $DistributionGroupMembers = Get-DistributionGroupMember -ResultSize Unlimited -Identity $DistributionGroup.Name
    $DistributionGroupsArray = $DistributionGroupsArray + [PSCustomObject]@{
        Name                                                   = $DistributionGroup.Name ;
        DisplayName                                            = $DistributionGroup.DisplayName ;
        ManagedBy                                              = $DistributionGroupMembers.ManagedBy -join ";" ;
        GroupType                                              = $DistributionGroup.GroupType ;
        RecipientTypeDetails                                   = $distributionGroup.RecipientTypeDetails ;
        Id                                                     = $distributionGroup.Id ;
        PrimarySmtpAddress                                     = $distributionGroup.PrimarySmtpAddress -join ";" ;
        WindowsEmailAddress                                    = $distributionGroup.WindowsEmailAddress -join ";" ;
        EmailAddresses                                         = $distributionGroup.EmailAddresses -join ";" ;
        HiddenFromAddressListsEnabled                          = $distributionGroup.HiddenFromAddressListsEnabled ;
        MaxSendSize                                            = $distributionGroup.MaxSendSize ;
        MaxReceiveSize                                         = $distributionGroup.MaxReceiveSize ;
        ModeratedBy                                            = $distributionGroup.ModeratedBy -join ";" ;
        IsDirSynced                                            = $distributionGroup.IsDirSynced ;
        GrantSendOnBehalfTo                                    = $distributionGroup.GrantSendOnBehalfTo -join ";" ;
        AcceptMessagesOnlyFromDLMembersWithDisplayNames        = $distributionGroup.AcceptMessagesOnlyFromDLMembersWithDisplayNames -join ";" ;
        AcceptMessagesOnlyFromSendersOrMembersWithDisplayNames = $distributionGroup.AcceptMessagesOnlyFromSendersOrMembersWithDisplayNames -join ";" ;
        AcceptMessagesOnlyFromWithDisplayNames                 = $distributionGroup.AcceptMessagesOnlyFromWithDisplayNames -join ";" ;
        BccBlocked                                             = $distributionGroup.BccBlocked ;
        BypassModerationFromSendersOrMembersWithDisplayNames   = $distributionGroup.BypassModerationFromSendersOrMembersWithDisplayNames ;
        Description                                            = $distributionGroup.Description ;
        GrantSendOnBehalfToWithDisplayNames                    = $distributionGroup.GrantSendOnBehalfToWithDisplayNames -join ";"  ;
        HiddenGroupMembershipEnabled                           = $distributionGroup.HiddenGroupMembershipEnabled ;
        MemberDepartRestriction                                = $distributionGroup.MemberDepartRestriction ;
        MemberJoinRestriction                                  = $distributionGroup.MemberJoinRestriction ;
        MigrationToUnifiedGroupInProgress                      = $distributionGroup.MigrationToUnifiedGroupInProgress ;
        ModeratedByWithDisplayNames                            = $distributionGroup.ModeratedByWithDisplayNames -join ";"  ;
        ModerationEnabled                                      = $distributionGroup.ModerationEnabled ;
        RejectMessagesFrom                                     = $distributionGroup.RejectMessagesFrom -join ";" ;
        RejectMessagesFromDLMembers                            = $distributionGroup.RejectMessagesFromDLMembers -join ";"  ;
        RejectMessagesFromSendersOrMembers                     = $distributionGroup.RejectMessagesFromSendersOrMembers -join ";" ;
        RejectMessagesFromSendersOrMembersWithDisplayNames     = $distributionGroup.RejectMessagesFromSendersOrMembersWithDisplayNames -join ";"  ;
        ReportToManagerEnabled                                 = $distributionGroup.ReportToManagerEnabled ;
        ReportToOriginatorEnabled                              = $distributionGroup.ReportToOriginatorEnabled ;
        RequireSenderAuthenticationEnabled                     = $distributionGroup.RequireSenderAuthenticationEnabled ;
        SendOofMessageToOriginatorEnabled                      = $distributionGroup.SendOofMessageToOriginatorEnabled ;
        WhenChanged                                            = $distributionGroup.WhenChanged ;
        WhenCreated                                            = $distributionGroup.WhenCreated ;
        MemberCount                                            = $DistributionGroupMembers.Count ;
        ORGID                                                  = $OrgID
    }
    $i++
}

Write-Progress  -ID 1 -Activity "Processing Distribution Groups" -Completed
Start-Sleep 5
If ($DistributionGroupsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $DistributionGroupsArray | Export-Excel -Path $Output -AutoSize -TableName Groups_Distribution -WorksheetName Groups_Distribution  
    $DistributionGroupsArray | Export-CSV .\csvfiles\Groups_Distribution.CSV -NoTypeInformation  
    Start-Sleep 5
}
$P++
#endregion Distribution Groups

#region Dynamic Distribution Groups
$Process = "Dynamic Distribution Groups"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Dynamic Distribution Groups data 
$DynamicDistributionGroups = Get-DynamicDistributionGroup -ResultSize Unlimited | Sort-Object DisplayName
$DynamicDistributionGroupsArray = @()
$i = 1
Foreach ($DynamicDistributionGroup in $DynamicDistributionGroups) {
    If ($DynamicDistributionGroups.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Dynamic Distribution Groups" -Status "Dynamic Distribution Group $i of $($DynamicDistributionGroups.Count)" -PercentComplete (($i / $DynamicDistributionGroups.Count) * 100) 
    }
    $DynamicDistributionGroupsArray = $DynamicDistributionGroupsArray + [PSCustomObject]@{
        Name                 = $DynamicDistributionGroup.Name ;
        DisplayName          = $DynamicDistributionGroup.DisplayName ;
        RecipientFilterType  = $DynamicDistributionGroup.RecipientFilterType ;
        RecipientTypeDetails = $DynamicDistributionGroup.RecipientTypeDetails; 
        PrimarySmtpAddress   = $DynamicDistributionGroup.PrimarySmtpAddress;
        ManagedBy            = $DynamicDistributionGroup.ManagedBy -join ',' ;
        ORGID                = $OrgID
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If ($DynamicDistributionGroupsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $DynamicDistributionGroupsArray | Export-Excel -Path $Output -AutoSize -TableName Groups_DynamicDistribution -WorksheetName Groups_DynamicDistribution  
    $DynamicDistributionGroupsArray | Export-CSV .\csvfiles\Groups_DynamicDistribution.CSV -NoTypeInformation  
}
Write-Progress  -ID 1 -Activity "Processing Dynamic Distribution Groups" -Completed
$P++    
#endregion Dynamic Distribution Groups

#region Unified Group Mailboxes
$Process = "Unified Group Mailboxes"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$UnifiedGroupMailboxesArray = @()
$i = 1
foreach ($UnifiedGroup in $UnifiedGroups) {
    If ($UnifiedGroups.Count -gt "1") {
        Write-Progress  -ID 1  -Activity "Processing Unified Group Mailboxes" -Status "Unified Group Mailbox $i of $($UnifiedGroups.Count)" -PercentComplete (($i / $UnifiedGroups.Count) * 100)
    }
    $GroupName = $UnifiedGroup.Name
    $GroupID = $GroupName.Substring($GroupName.IndexOf('_') + 1)
    $Isteam = ""
    $Type = ""
    $IsTeam = $AllTeams | Where { $_.GroupID -eq $GroupID }
    If ($IsTeam -ne $Null) {
        $Type = "Team"
    }
    Else {
        $Type = "Group"
    }
    $UnifiedMailboxStats = Get-MailboxStatistics -Identity $UnifiedGroup.PrimarySMTPAddress | Select-Object LastLogonTime, DisplayName, @{Name = "TotalItemSizeMB"; Expression = { [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }, ItemCount, DeletedItemCount, @{Name = "TotalDeletedItemSizeMB"; Expression = { [math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }
    $UnifiedGroupMailboxesArray = $UnifiedGroupMailboxesArray + [PSCustomObject]@{
        Type                   = $Type ;
        DisplayName            = $UnifiedGroup.DisplayName ;
        Alias                  = $UnifiedGroup.Identity ;
        PrimarySmtpAddress     = $UnifiedGroup.PrimarySmtpAddress ;
        RecipientTypeDetails   = $UnifiedGroup.RecipientTypeDetails ;
        ItemCount              = $UnifiedMailboxStats.ItemCount; 
        TotalItemSizeMB        = $UnifiedMailboxStats.TotalItemSizeMB; 
        DeletedItemCount       = $UnifiedMailboxStats.DeletedItemCount;  
        TotalDeletedItemSizeMB = $UnifiedMailboxStats.TotalDeletedItemSizeMB;
        WhenCreatedUTC         = $UnifiedGroup.WhenCreatedUTC ;
        WhenChangedUTC         = $UnifiedGroup.WhenChangedUTC ;
        LastLogonTime          = $UnifiedMailboxStats.LastLogonTime ;
        EmailAddresses         = $UnifiedGroup.EmailAddresses -join ',';
        Guid                   = $UnifiedGroup.Guid ;
        ORGID                  = $OrgID
    }
    $i++
}
Start-Sleep 5
If ($UnifiedGroupMailboxesArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $UnifiedGroupMailboxesArray | Export-Excel -Path $Output -AutoSize -TableName Mailbox_UnifiedGroup -WorksheetName Mailbox_UnifiedGroup  
    $UnifiedGroupMailboxesArray | Export-CSV .\csvfiles\Mailbox_UnifiedGroup.CSV -NoTypeInformation  
}
Write-Progress  -ID 1  -Activity "Processing Unified Group Mailboxes" -Completed    
$P++
#endregion 

#region Unified Groups
$Process = "Unified Groups"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Microsoft 365 Groups data
$UnifiedGroupArray = @()
$i = 1
Foreach ($UnifiedGroup in $UnifiedGroups) {
    If ($UnifiedGroups.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Unified Groups" -Status "Unified Group $i of $($UnifiedGroups.Count)" -PercentComplete (($i / $UnifiedGroups.Count) * 100)  
    }
    $GroupName = $UnifiedGroup.Name
    $GroupID = $UnifiedGroup.ExternalDirectoryObjectId 
    $Isteam = ""
    $Type = ""
    $IsTeam = $AllTeams | Where { $_.GroupID -eq $GroupID }
    If ($IsTeam -ne $Null) {
        $Type = "Team"
    }
    Else {
        $Type = "Group"
    }

    $SharePointSite = $SPOSites | Where { $_.URL -eq $UnifiedGroup.SharePointSiteUrl }

    $GroupMailboxStats = $UnifiedGroupMailboxesArray | Where { $_.PrimarySMTPAddress -eq $UnifiedGroup.PrimarySMTPAddress }

    $UnifiedGroupArray = $UnifiedGroupArray + [PSCustomObject]@{
        Type                     = $Type ;
        Name                     = $UnifiedGroup.Name ;
        DisplayName              = $UnifiedGroup.DisplayName ;
        AccessType               = $UnifiedGroup.AccessType ;
        PrimarySMTPAddress       = $UnifiedGroup.PrimarySMTPAddress ;
        MailboxTotalItemSizeMB   = $GroupMailboxStats.TotalItemSizeMB ;
        SharePointStorageMB      = $SharePointSite.StorageUsageCurrent ;
        MemberCount              = $UnifiedGroup.GroupMemberCount  ;
        GroupExternalMemberCount = $UnifiedGroup.GroupExternalMemberCount
        SharePointSiteURL        = $UnifiedGroup.SharePointSiteUrl ;
        SharePointDocumentsUrl   = $UnifiedGroup.SharePointDocumentsUrl ;
        SharePointNotebookUrl    = $UnifiedGroup.SharePointNotebookUrl ; 
        GroupID                  = $GroupID ;
        ORGID                    = $OrgID
    }
    $i++
}  
Start-Sleep 5
If ($UnifiedGroupArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $UnifiedGroupArray | Export-Excel -Path $Output -AutoSize -TableName Groups_Unified -WorksheetName Groups_Unified  
    $UnifiedGroupArray | Export-CSV .\csvfiles\Groups_Unified.csv -NoTypeInformation  
}
Write-Progress  -ID 1 -Activity "Processing Unified Groups" -Completed
$P++
#endregion Unified Groups

#region Mailboxes
$Process = "Mailboxes"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$AllMailboxesArray = @()
$InactivePeriod = (Get-Date).AddDays(-90)
$i = 1
Foreach ($Mailbox in $AllMailboxes) {
    If ($AllMailboxes.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Mailbox Data" -Status "Mailbox $i of $($AllMailboxes.Count)" -PercentComplete (($i / $AllMailboxes.Count) * 100)  
    }
    $MailboxStats = Get-MailboxStatistics -Identity $Mailbox.UserPrincipalName | Select-Object LastLogonTime, DisplayName, @{Name = "TotalItemSizeMB"; Expression = { [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }, ItemCount, DeletedItemCount, @{Name = "TotalDeletedItemSizeMB"; Expression = { [math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }
    If ($Mailbox.ArchiveStatus -eq "Active") {
        $ArchiveStats = Get-MailboxStatistics -Identity $Mailbox.UserPrincipalName -Archive | Select-Object DisplayName, @{Name = "TotalItemSizeMB"; Expression = { [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }, ItemCount, DeletedItemCount, @{Name = "TotalDeletedItemSizeMB"; Expression = { [math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }
    }
    Else {
        $ArchiveStats = ""
    }
    $AllMailboxesArray = $AllMailboxesArray + [PSCustomObject]@{
        Type                          = $Mailbox.RecipientTypeDetails ;
        UserPrincipalName             = $Mailbox.UserPrincipalName ;
        DisplayName                   = $Mailbox.DisplayName ;
        Alias                         = $Mailbox.Identity ;
        PrimarySmtpAddress            = $Mailbox.PrimarySmtpAddress ;
        ItemCount                     = $MailboxStats.ItemCount; 
        TotalItemSizeMB               = $MailboxStats.TotalItemSizeMB; 
        DeletedItemCount              = $MailboxStats.DeletedItemCount;  
        TotalDeletedItemSizeMB        = $MailboxStats.TotalDeletedItemSizeMB;
        Archive                       = $Mailbox.ArchiveStatus ;
        ArchiveDisplayName            = $ArchiveStats.DisplayName
        ArchiveItemCount              = $ArchiveStats.ItemCount ; 
        ArchiveTotalItemSizeMB        = $ArchiveStats.TotalItemSizeMB ; 
        ArchiveDeletedItemCount       = $ArchiveStats.DeletedItemCount ; 
        ArchiveTotalDeletedItemSizeMB = $ArchiveStats.TotalDeletedItemSizeMB ; 
        WhenCreatedUTC                = $Mailbox.WhenCreatedUTC ;
        WhenChangedUTC                = $Mailbox.WhenChangedUTC ;
        LastLogonTime                 = $MailboxStats.LastLogonTime ;
        EmailAddresses                = $Mailbox.EmailAddresses -join ',';
        LitigationHoldEnabled         = $Mailbox.LitigationHoldEnabled  ;
        LitigationHoldDuration        = $Mailbox.LitigationHoldDuration ;
        InPlaceHolds                  = $Mailbox.InPlaceHolds -join ',' ;
        RetentionPolicy               = $Mailbox.RetentionPolicy ;
        RetentionHoldEnabled          = $Mailbox.RetentionHoldEnabled ;
        StartDateForRetentionHold     = $Mailbox.StartDateForRetentionHold ; 
        EndDateForRetentionHold       = $Mailbox.EndDateForRetentionHold ;
        AccountSKU                    = $licenseString -join ',' ;
        Guid                          = $Mailbox.Guid ;
        ORGID                         = $OrgID
    }
    $i++
}

If ($AllMailboxesArray -ne $Null) {
    $AllMailboxesArray | Export-Excel -Path $Output -AutoSize -TableName Mailbox_Account -WorksheetName Mailbox_Account 
    $AllMailboxesArray | Export-CSV .\csvfiles\Mailbox_Account.csv -NoTypeInformation 
} 
Start-Sleep 5
$Inactivemailboxes = $AllMailboxesArray | Where-Object { $_.Lastlogontime -lt (Get-Date).AddDays(-90) }
If ($Inactivemailboxes -ne $Null) {
    $Inactivemailboxes | Export-Excel -Path $Output -AutoSize -TableName Mailbox_Inactive -WorksheetName Mailbox_Inactive  
}
Start-Sleep 5
$LargetsMailboxes = $AllMailboxesArray | Sort-Object TotalItemSizeMB -Descending | Select-Object -First 10
If ($LargetsMailboxes -ne $Null) {
    $LargetsMailboxes | Export-Excel -Path $Output -AutoSize -TableName Mailbox_TopTen -WorksheetName Mailbox_TopTen  
}
Write-Host "Writing $Process data to $Output"
Write-Progress -ID 1 -Activity "Gathering Mailbox Data" -Completed
$P++
#endregion Mailboxes

#region Public Folders
$Process = "Public Folders"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Public Folder data
$ErrorActionPreference = "SilentlyContinue"
$PublicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited | Sort-Object Parentpath 
$ErrorActionPreference = "Continue"
$PublicFolderArray = @()
$i = 1
Foreach ($PublicFolder in $PublicFolders) {
    If ($PublicFolders.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Public Folders" -Status "Public Folder $i of $($PublicFolders.Count)" -PercentComplete (($i / $PublicFolders.Count) * 100)  
    }
    $PublicFolderStats = Get-PublicFolderStatistics -Identity $PublicFolder.Identity | Select-Object Name, @{Name = "TotalItemSizeMB"; Expression = { [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }, ItemCount, DeletedItemCount, @{Name = "TotalDeletedItemSizeMB"; Expression = { [math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }
    $PublicFolderArray = $PublicFolderArray + [PSCustomObject]@{
        Name            = $PublicFolder.Name ;
        Identity        = $PublicFolder.Identity ;
        ParentPath      = $PublicFolder.ParentPath ;
        ItemCount       = $PublicFolderStats.ItemCount; 
        TotalItemSizeMB = $PublicFolderStats.TotalItemSizeMB; 
        FolderClass     = $PublicFolder.FolderClass ;
        MailEnabled     = $PublicFolder.MailEnabled ;
        ORGID           = $OrgID 
    }
    $i++
}
Start-Sleep 5
If ($PublicFolderArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $PublicFolderArray | Export-Excel -Path $Output -AutoSize -TableName PublicFolders_All -WorksheetName PublicFolders_All
    $PublicFolderArray | Export-CSV .\csvfiles\PublicFolders_All.csv -NoTypeInformation
}
Write-Progress  -ID 1 -Activity "Processing Public Folders" -Completed
$P++
#endregion Public Folders

#region SharePoint Sites
$Process = "SharePoint Sites"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets SharePoint site data
$SPOSitesArray = @()
$SPOTeamsChannelsArray = @()
$i = 1
ForEach ($SPOSite in $SPOSites) {
    If ($SPOSites.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing SharePoint Sites" -Status "Site $i of $($SPOSites.Count)" -PercentComplete (($i / $SPOSites.Count) * 100) 
    }
    $Type = ""
    $SPOUnifiedGroup = ""
    $SPOUnifiedGroup = $UnifiedGroupArray | Where { $_.SharePointSiteURL -eq $SPOSite.URL }

    If ($SPOUnifiedGroup.Type -eq "Team") {
        $Type = "Team"
    }
    ElseIf ($SPOUnifiedGroup.Type -eq "Group") {
        $Type = "Group"
    }
    ElseIf ($SPOSite.Template -Like "TeamChannel*") {
        $Type = "Channel"
    }
    Else {
        $Type = "SPO"
    }


    $SPOSitesArray = $SPOSitesArray + [PSCustomObject]@{
        Type                                     = $Type ; 
        Title                                    = $SPOSite.Title ; 
        LocaleId                                 = $SPOSite.LocaleId ; 
        Url                                      = $SPOSite.Url ; 
        Status                                   = $SPOSite.Status ;
        LastContentModifiedDate                  = $SPOSite.LastContentModifiedDate ; 
        Owner                                    = $SPOSite.Owner ; 
        Template                                 = $SPOSite.Template ; 
        ResourceUsageCurrent                     = $SPOSite.ResourceUsageCurrent ; 
        ResourceUsageAverage                     = $SPOSite.ResourceUsageAverage ; 
        StorageUsageCurrentMB                    = $SPOSite.StorageUsageCurrent ; 
        ConditionalAccessPolicy                  = $SPOSite.ConditionalAccessPolicy ;
        SensitivityLabel                         = $SPOSite.SensitivityLabel ;
        AllowSelfServiceUpgrade                  = $SPOSite.AllowSelfServiceUpgrade ;
        AllowEditing                             = $SPOSite.AllowEditing ; 
        SharingAllowedDomainList                 = $SPOSite.SharingAllowedDomainList ; 
        SharingBlockedDomainList                 = $SPOSite.SharingBlockedDomainList ; 
        DenyAddAndCustomizePages                 = $SPOSite.DenyAddAndCustomizePages ;
        BlockDownloadLinksFileType               = $SPOSite.BlockDownloadLinksFileType ;
        DefaultLinkPermission                    = $SPOSite.DefaultLinkPermission ;
        DefaultSharingLinkType                   = $SPOSite.DefaultSharingLinkType ; 
        DisableAppViews                          = $SPOSite.DisableAppViews ; 
        DisableCompanyWideSharingLinks           = $SPOSite.DisableCompanyWideSharingLinks ; 
        DisableFlows                             = $SPOSite.DisableFlows ; 
        LimitedAccessFileType                    = $SPOSite.LimitedAccessFileType ; 
        LockState                                = $SPOSite.LockState ; 
        SandboxedCodeActivationCapability        = $SPOSite.SandboxedCodeActivationCapability ; 
        SharingCapability                        = $SPOSite.SharingCapability ; 
        ShowPeoplePickerSuggestionsForGuestUsers = $SPOSite.ShowPeoplePickerSuggestionsForGuestUsers ; 
        SharingDomainRestrictionMode             = $SPOSite.SharingDomainRestrictionMode ; 
        LockIssue                                = $SPOSite.LockIssue ; 
        WebsCount                                = $SPOSite.WebsCount ; 
        CompatibilityLevel                       = $SPOSite.CompatibilityLevel ; 
        DisableSharingForNonOwnersStatus         = $SPOSite.DisableSharingForNonOwnersStatus ; 
        HubSiteId                                = $SPOSite.HubSiteId ; 
        IsHubSite                                = $SPOSite.IsHubSite ; 
        RelatedGroupId                           = $SPOSite.RelatedGroupId ; 
        GroupId                                  = $SPOSite.GroupId ;
        ORGID                                    = $OrgID 
    }  
    $i++
}
$SPOTemplates = $SPOSites.Template
$SPOTemplatesGroup = $SPOTemplates | Group-Object
Start-Sleep 5
If ($SPOSitesArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    #All Sites
    $SPOSitesArray | Export-Excel -Path $Output -AutoSize -TableName SPOSites_All -WorksheetName SPOSites_All 
    Start-Sleep 5 
    #Largest Sites
    $LargestSites = $SPOSitesArray | Sort-Object StorageUsageCurrentMB -Descending | Select-Object -First 10
    $LargestSites | Export-Excel -Path $Output -AutoSize -TableName SPOSites_TopTen -WorksheetName SPOSites_TopTen
    Start-Sleep 5
    #Microsoft Teams Channels
    $SPOTeamsChannelsArray = $SPOSitesArray | Where-Object { $_.M365Group -ne "Yes" -and $_.Template -Like "TEAMCHANNEL*" } 
    $SPOTeamsChannelsArray | Export-Excel -Path $Output -AutoSize -TableName SPOSites_TeamsChannels -WorksheetName SPOSites_TeamsChannels
    Start-Sleep 5
    $SPOSitesArray | Export-CSV .\csvfiles\SPOSites_All.csv -NoTypeInformation 

}
Write-Progress  -ID 1 -Activity "Processing SharePoint Sites" -Completed   
$P++
#endregion SharePoint Sites

#region Microsoft Teams
$Process = "Microsoft Teams"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$TeamArray = @()
$i = 1
Foreach ($Team in $AllTeams) {  
    If ($AllTeams.Count -gt "1") {
        Write-Progress  -Activity "Processing Microsoft Teams" -Status "Microsoft Team $i of $($AllTeams.Count)" -PercentComplete (($i / $AllTeams.Count) * 100) 
    }
    $TeamGroup = $UnifiedGroupArray | Where { $_.GroupID -eq $Team.GroupID }
    $SharePointSite = $SPOSitesArray | Where { $_.URL -eq $TeamGroup.SharePointSiteURL } #Get-SPOSite $TeamGroup.SharePointSiteUrl
    $Channels = Get-TeamChannel -GroupId $Team.GroupId
    $ChannelsCount = $Channels | Measure-Object
    $TotalUsers = [int]$TeamOwnerCount.Count + [int]$TeamMemberCount.Count + [int]$TeamGuestCount.Count
    $TeamArray = $TeamArray + [PSCustomObject]@{
        GroupId             = $Team.GroupId ; 
        DisplayName         = $Team.DisplayName ; 
        Description         = $Team.Description ; 
        Visibility          = $Team.Visibility ; 
        MailNickName        = $Team.MailNickName ; 
        Classification      = $Team.Classification ; 
        Archived            = $Team.Archived ; 
        StorageMB           = $SharePointSite.StorageUsageCurrentMB ;
        Channels            = $ChannelsCount.Count ; 
        OwnerCount          = $TeamGroup.OwnerCount ; 
        MemberCount         = $TeamGroup.MemberCount  ;
        ExternalMemberCount = $TeamGroup.GroupExternalMemberCount  ;
        Owners              = $TeamGroup.Owners ; 
        Members             = $TeamGroup.Members ;
        ORGID               = $OrgID
    }
    $i++
}
Start-Sleep 5
If ($TeamArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"    
    $TeamArray | Export-Excel -Path $Output -AutoSize -TableName Teams_All -WorksheetName Teams_All  
    $TeamArray | Export-CSV .\csvfiles\Teams_All.CSV -NoTypeInformation  
}
Write-Progress  -ID 1 -Activity "Processing Microsoft Teams" -Completed
$P++

#endregion Microsoft Teams

#region Teams Channels
$Process = "Microsoft Teams Channels"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$TeamWithChannels  = @()
$TeamChannelArray = @()

$TeamWithChannels = $TeamArray | Where { $_.Channels -gt "1" }
$i = 1
Foreach ($TeamWithChannel in $TeamWithChannels) {  
    If ($Channel.Count -gt "1") {
        Write-Progress  -Activity "Processing Microsoft Teams Channels" -Status "Microsoft Team $i of $($AllTeams.Count)" -PercentComplete (($i / $AllTeams.Count) * 100) 
    }
    $Channels = Get-TeamChannel -GroupId $TeamWithChannel.Groupid
    Foreach ($Channel in $Channels) {
        If($Channel.Displayname -ne "General") {
            $TeamChannelArray = $TeamChannelArray + [PSCustomObject]@{
                GroupId            = $TeamWithChannel.GroupId ; 
                TeamDisplayName    = $TeamWithChannel.DisplayName ;
                ChannelID          = $Channel.ID ;
                ChannelDisplayname = $Channel.Displayname ;
                Description        = $Channel.Description ;
                MembershipType     = $Channel.membershipType ;
                ORGID              = $OrgID
            }
        }
    }
    $i++
}
Start-Sleep 5
If ($TeamChannelArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"    
    $TeamChannelArray | Export-Excel -Path $Output -AutoSize -TableName Teams_Channels -WorksheetName Teams_Channels 
    $TeamChannelArray | Export-CSV .\csvfiles\Teams_Channels.csv -NoTypeInformation  
 
}
Write-Progress  -ID 1 -Activity "Processing Microsoft Teams Channels" -Completed
$P++
#endregion Teams Channels

#region Teams Usage Reports
$Process = "Teams Usage Reports"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$MgReportTeamUserActivityUserDetailCSV = Get-MgReportTeamUserActivityUserDetail -Period "D180" -OutFile TeamsUserActivity_Last180Days.csv
$MgReportTeamUserActivityUserDetails = Import-csv TeamsUserActivity_Last180Days.csv
$MgReportTeamUserActivityUserArray = @()
$i = 1
Foreach ($MgReportTeamUserActivityUserDetail in $MgReportTeamUserActivityUserDetails) {
    If ($MgReportTeamUserActivityUserDetails.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Teams User Activity" -Status "User $i of $($MgReportTeamUserActivityUserDetails.Count)" -PercentComplete (($i / $MgReportTeamUserActivityUserDetails.Count) * 100) 
    }
    $MgReportTeamUserActivityUserArray = $MgReportTeamUserActivityUserArray + [PSCustomObject]@{
        UserPrincipalname  = $MgReportTeamUserActivityUserDetail."User Principal Name" ;
        LastActivityDate     = $MgReportTeamUserActivityUserDetail."Last Activity Date " ;
        PrivateChatMessageCount     = $MgReportTeamUserActivityUserDetail."Private Chat Message Count"  ;
        PostMessages     = $MgReportTeamUserActivityUserDetail."Post Messages"  ;
        CallCount     = $MgReportTeamUserActivityUserDetail."Call Count"  ;
        MeetingCount      = $MgReportTeamUserActivityUserDetail."Meeting Count "  ;
        IsDeleted     = $MgReportTeamUserActivityUserDetail."Is Deleted"  ;
        ORGID = $OrgID
    }
    $i++
}
Start-Sleep 5
If ($MgReportTeamUserActivityUserArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MgReportTeamUserActivityUserArray | Export-Excel -Path $Output -AutoSize -TableName Teams_User_Activity -WorksheetName Teams_User_Activity  
}
Write-Progress  -ID 1 -Activity "Teams Usage Reports" -Completed
$P++
#endregion Teams Usage Reports

#region MTO Status
$Process = "MTO Status"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$MTOStatus = @()
$MTOStatus = Get-MgTenantRelationshipMultiTenantOrganization
$MTOStatusArray = @()
$i = 1
Foreach ($MTO in $MTOStatus) {
    If ($MTOStatus.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing MTO Status" -Status "MTO $i of $($MTOStatus.Count)" -PercentComplete (($i / $MTOStatus.Count) * 100) 
    }
    $MTOStatusArray = $MTOStatusArray + [PSCustomObject]@{
        ID                  = $MTO.ID ;
        DisplayName         = $MTO.DisplayName ;
        Description         = $MTO.Description ;
        State               = $MTO.State ;
    }
    $i++
}
Start-Sleep 5
Write-Progress  -ID 1 -Activity "MTO Status" -Completed
$P++
#endregion MTO Status


#region Finishing Up
$Process = "Miscellaneous Data"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Adds data to the Dashboard or cells
$CompanyName = $Org.Displayname 
$ReportDate = (Get-Date).ToString()
If ($Org.OnPremisesSyncEnabled -eq $Null) {
    $EntraConnectStatus = "Disabled"
}
If ($Org.OnPremisesSyncEnabled -eq $True) {
    $EntraConnectStatus = "Enabled"
}
If ($Org.OnPremisesSyncEnabled -eq $False) {
    $EntraConnectStatus = "Broken"
}
$EXOHybrid = Get-OnPremisesOrganization
If ($EXOHybrid -eq $Null) {
    $ExchangeHybridStatus = "Disabled"
}
If ($EXOHybrid.Isvalid -eq $True) {
    $ExchangeHybridStatus = "Enabled"
}
$EndTime = $(get-date) - $StartTime
$TotalTime = "{0:HH:mm:ss}" -f ([datetime]$EndTime.Ticks)
$DefaultDomainName = $AcceptedDomains | Where { $_.Default -eq $True }

$OnMicrosoftDomain = $AcceptedDomains | Where-Object { $_.DomainName -like "*.onmicrosoft.com" } | Select-Object DomainName -ExpandProperty DomainName
$OnMicrosoftPrefix = $OnMicrosoftDomain.split('.')[0] 
$AdminURL = "https://" + $OnMicrosoftPrefix + "-admin.sharepoint.com"

Start-Sleep 5
Write-Host "Writing $Process data to $Output"
$Excel = Open-ExcelPackage -Path $Output
$worksheet = $excel.Workbook.Worksheets['Data']
$worksheet.Cells['E9'].value = $Output
$worksheet.Cells['E10'].value = $ReportDate
$worksheet.Cells['E11'].value = $ScriptVersion
$worksheet.Cells['E12'].value = $TemplateVersion
$worksheet.Cells['E13'].value = $TotalTime
$worksheet.Cells['E14'].value = $Consultant
$worksheet.Cells['E18'].value = $CompanyName
$worksheet.Cells['E19'].value = $Org.ID
$worksheet.Cells['E20'].value = $AdminURL
$worksheet.Cells['E21'].value = $DefaultDomainName.Name
#$worksheet.Cells['D50'].value = $InactivePeriod
$worksheet.Cells['E25'].value = $EntraConnectStatus
$worksheet.Cells['E26'].value = $ExchangeHybridStatus
$worksheet.Cells['P56'].value = $MTO.ID
$worksheet.Cells['P57'].value = $MTO.State
$worksheet.Cells['P58'].value = $MTO.DisplayName
$worksheet.Cells['P59'].value = $MTO.Description
Close-ExcelPackage $Excel
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -Completed
$P++

$Process = "Cleaning Up"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
Write-Host "Data written to $Output"
Write-Progress -ID 3 -Activity "Having a rest before opening the report in Excel"
Start-Sleep -Seconds 10  
Stop-Transcript

$continue =""
$continue = Read-Host "Do you want to open the report? (Y/N)"
if ($continue -ne "Y") {
    Write-Host "Script execution has ended." -ForegroundColor Red
    Break
}
if ($continue -eq "Y") {
    Export-Excel $Output -show
}

#endregion

#endregion script