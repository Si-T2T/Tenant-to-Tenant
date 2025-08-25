$Script = "MGGraph"
$Version = "v2.0.0"
Write-Host "Running" $Script $Version -ForegroundColor Green

Disconnect-MgGraph -ErrorAction SilentlyContinue
Disconnect-msgraph -ErrorAction SilentlyContinue
Write-Host "Connecting to Microsoft Graph" -ForegroundColor Green
$Scope = @{
    Scopes = @(
        "Application.Read.All",
        "Auditlog.Read.All",
        "BitlockerKey.ReadBasic.All",
        "ChannelMessage.Read.All",
        "Device.Read.All",
        "Directory.Read.All",
        "Files.Read.All",
        "Group.Read.All",
        "OrgContact.Read.All",
        "Organization.Read.All",        
        "Policy.Read.All",
        "Reports.Read.All",
        "Sites.Read.All",
        "User.Read.All"
    )
}
        

Connect-MGGraph @Scope -nowelcome
$Properties = @()
$Properties = @('ID', 'UserPrincipalName', 'DisplayName', 'Mail', 'LicenseAssignmentStates', 'UsageLocation', 'UserType', 'AccountEnabled', 'OnPremisesSyncEnabled')
$MGUsers = Get-MgUser -All -Property $Properties | Sort-Object displayname
Sleep 10

$MGUsersArray = @()
$i = 1
Foreach ($MGUser in $MGUsers) {
    If ($MGUsers.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing All MG User Accounts" -Status "User Account $i of $($MGUsers.Count)" -PercentComplete (($i / $MGUsers.Count) * 100) 
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
    }
    $i++
}
Start-Sleep 5
    Write-Host "Writing MGUser Data"    
    $MGUsersArray | Export-csv  .\csvfiles\AllMGUsers.csv -notypeinformation


Write-Host "User data exported"

# Get Friendly License Names
$MSLicenseInfo = Invoke-WebRequest -UseBasicParsing "https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference"
    $SkuCSVLink=($MSLicenseInfo.links | Where-Object {$_.href -match ".csv"}).href
    $WebClient=[System.Net.WebClient]::New()
    $InMemoryCSV = $WebClient.DownloadString($SkuCSVLink)
    $ProductSkuCSV = ConvertFrom-CSV -Delimiter ',' -InputObject $InMemoryCSV
    $WebClient.Dispose()

# Get all tenant skus
[Array]$Skus = Get-MgSubscribedSku
$TenantLicenseDetails = $Skus | select SkuPartNumber, ConsumedUnits, @{ n = 'TotalUnits'; e = { $_.prepaidunits.enabled } }, @{ n = 'FriendlyName'; e= {$_ | foreach { $FriendlyLicenses[$_.SkuPartNumber] } } }
[Array]$Users = Get-MGUser -All
$i = 0
foreach ($user in $Users)
{
	Write-Progress -Activity "Processing User License details" -Status "Working on $($user.displayname)" -PercentComplete (($i / $Users.Count) * 100)
	$user.LicenseDetails = Get-MgUserLicenseDetail -UserId $user.id
	$i++
}
$UserLicenseDetails = $Users | where LicenseDetails | select UserPrincipalName, @{ n = 'Licenses'
e = { 
		($_ | foreach { $_.licensedetails | foreach {
			$a = $_
			$FriendlyLicense = ($ProductSkuCSV | Where-Object {$_.String_ID -eq $a.SkuPartNumber} | Select-Object -First 1).Product_Display_Name
			if ($FriendlyLicense) { $FriendlyLicense }
					else { $a.SkuPartID }
				}
			}) -join ';'
	}
}

Start-Sleep 5
    Write-Host "Writing License Data"    
    $MGUsersArray | Export-csv  .\csvfiles\LicenseData.csv -notypeinformation


Write-Host "User data exported"
Pause