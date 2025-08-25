$Script = "ProcessData"
$Version = "v2.0.0"
Write-Host "Running" $Script $Version -ForegroundColor Green

$Consultant = Get-Content .\csvfiles\Consultant.txt

#region Process Data

[Array]$AllMGUsers = Import-Csv .\csvfiles\AllMGUsers.csv
[Array]$AllMailboxes = Import-Csv .\csvfiles\AccountMailboxes.csv
[Array]$OneDrives = Import-Csv .\csvfiles\OneDrives.csv
[Array]$SPOSites = Import-Csv .\csvfiles\SPOSites.csv
[Array]$Teams = Import-Csv .\csvfiles\Teams.csv
[Array]$TeamsChannels = Import-Csv .\csvfiles\TeamChannels.csv
[Array]$MailboxUnifiedGroup = Import-Csv .\csvfiles\GroupMailboxs.csv

#region User Overview
$MGUsersOverviewArray = @()
Foreach($MGUser in $AllMGUsers) {
    $MailboxStats = ""
    $mailboxstats = $AllMailboxes | Where{$_.UserPrincipalName -eq $mguser.userprincipalname}
        $OneDriveStats = ""
    If($MGUser.LicenseSKUs -ne $Null) {
        $OneDriveStats = $OneDrives | Where{$_.Owner -eq $mguser.userprincipalname}
        $OneDriveGB = $OneDriveStats.StorageUsageCurrentGB
    }
     If($MGUser.LicenseSKUs -eq $Null) {
        $OneDriveGB = "-"
    }
    $MGUsersOverviewArray = $MGUsersOverviewArray + [PSCustomObject]@{
        ID                 = $MGUser.ID;
        DisplayName        = $MGUser.DisplayName ;
        UserPrincipalName  = $MGUser.UserPrincipalName ;
        Mail               = $MGUser.Mail ;
        UserType           = $MGUser.UserType ;
        UsageLocation      = $MGUser.UsageLocation ;
        AccountEnabled     = $MGUser.AccountEnabled ;
        LicenseSKUs        = $MGUser.LicenseSKUs -join ";" ;
        IsDirSynced        = $MGUser.IsDirSynced
        OneDriveGB         = $OneDriveGB

    }
}
$MGUsersOverviewArray

#endregion User Overview







#endregion Process Data

#region Process Data

[Array]$AllMGUsers = Import-Csv .\csvfiles\AllMGUsers.csv
[Array]$AllMailboxes = Import-Csv .\csvfiles\AccountMailboxes.csv
[Array]$OneDrives = Import-Csv .\csvfiles\OneDrives.csv
[Array]$SPOSites = Import-Csv .\csvfiles\SPOSites.csv
[Array]$Teams = Import-Csv .\csvfiles\Teams.csv
[Array]$TeamsChannels = Import-Csv .\csvfiles\TeamChannels.csv
[Array]$MailboxUnifiedGroup = Import-Csv .\csvfiles\GroupMailboxs.csv

#region User Overview
$MGUsersOverviewArray = @()
Foreach($MGUser in $AllMGUsers) {
    $MailboxStats = ""
    $mailboxstats = $AllMailboxes | Where{$_.UserPrincipalName -eq $mguser.userprincipalname}
        $OneDriveStats = ""
    If($MGUser.LicenseSKUs -ne $Null) {
        $OneDriveStats = $OneDrives | Where{$_.Owner -eq $mguser.userprincipalname}
        $OneDriveGB = $OneDriveStats.StorageUsageCurrentGB
    }
     If($MGUser.LicenseSKUs -eq $Null) {
        $OneDriveGB = "-"
    }
    $MGUsersOverviewArray = $MGUsersOverviewArray + [PSCustomObject]@{
        ID                 = $MGUser.ID;
        DisplayName        = $MGUser.DisplayName ;
        UserPrincipalName  = $MGUser.UserPrincipalName ;
        Mail               = $MGUser.Mail ;
        UserType           = $MGUser.UserType ;
        UsageLocation      = $MGUser.UsageLocation ;
        AccountEnabled     = $MGUser.AccountEnabled ;
        LicenseSKUs        = $MGUser.LicenseSKUs -join ";" ;
        IsDirSynced        = $MGUser.IsDirSynced
        OneDriveGB         = $OneDriveGB

    }
}


#endregion User Overview







#endregion Process Data