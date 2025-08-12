$Script = "SPO"
$Version = "v2.0.0"
Write-Host "Running" $Script $Version -ForegroundColor Green

Connect-ExchangeOnline -DisableWAM

$OnMicrosoftDomain = Get-AcceptedDomain | Where-Object { $_.DomainName -like "*.onmicrosoft.com" } | Select-Object DomainName -ExpandProperty DomainName
$OnMicrosoftPrefix = $OnMicrosoftDomain.split('.')[0] 
$AdminURL = "https://" + $OnMicrosoftPrefix + "-admin.sharepoint.com"
Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell
Connect-SPOService -URL $AdminURL



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

    Write-Host "Writing OneDrive data"
    $OneDriveArray | Export-CSV .\csvfiles\OneDrives.csv -NoTypeInformation           


$SPOSites = @()

$SPOSites = Get-SPOSite -IncludePersonalSite $False -Limit All | Where-Object{$_.URL -NotLike "*my.sharepoint.com/" -and $_.URL -notlike "*sharepoint.com/search"} | Sort-Object Title

$SPOSitesArray = @()
$SPOTeamsChannelsArray = @()
$i = 1
ForEach ($SPOSite in $SPOSites) {
    If ($SPOSites.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing SharePoint Sites" -Status "Site $i of $($SPOSites.Count)" -PercentComplete (($i / $SPOSites.Count) * 100) 
    }
    $SPOSitesArray = $SPOSitesArray + [PSCustomObject]@{
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

    Write-Host "Writing SharePoint Site data"
    $SPOSitesArray | Export-CSV .\csvfiles\SPOSites.csv -NoTypeInformation 






Pause