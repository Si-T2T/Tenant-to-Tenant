$Script = "UnifiedGroups"
$Version = "v2.0.0"
Write-Host "Running" $Script $Version -ForegroundColor Green

Connect-ExchangeOnline -DisableWAM

Write-Host "Getting Microsoft 365 Groups" -ForegroundColor Green

$UnifiedGroups = @()
$UnifiedGroups = Get-UnifiedGroup -ResultSize Unlimited | Sort-Object DisplayName

$UnifiedGroupArray = @()
$i = 1
foreach ($UnifiedGroup in $UnifiedGroups) {
    If ($UnifiedGroups.Count -gt "1") {
        Write-Progress  -ID 1  -Activity "Processing Unified Group Mailboxes" -Status "Unified Group Mailbox $i of $($UnifiedGroups.Count)" -PercentComplete (($i / $UnifiedGroups.Count) * 100)
    }
    $UnifiedGroupArray = $UnifiedGroupArray + [PSCustomObject]@{
        DisplayName            = $UnifiedGroup.DisplayName ;
        AccessType             = $UnifiedGroup.AccessType ;
        PrimarySmtpAddress     = $UnifiedGroup.PrimarySmtpAddress ;
        GroupID                = $UnifiedGroup.ExternalDirectoryObjectId ;
        SharePointSiteURL      = $UnifiedGroup.SharePointSiteURL 
    }
    $i++
}
Start-Sleep 5

    Write-Host "Writing Unified Group data"
    $UnifiedGroupArray | Export-CSV .\csvfiles\UnifiedGroups.CSV -NoTypeInformation  


Write-Host "Unified Group Data Exported" -ForegroundColor Green
Write-Host "Getting Microsoft 365 Group Mailbox Data" -ForegroundColor Green

$UnifiedGroupMailboxesArray = @()
$i = 1
foreach ($UnifiedGroup in $UnifiedGroups) {
    If ($UnifiedGroups.Count -gt "1") {
        Write-Progress  -ID 1  -Activity "Processing Unified Group Mailboxes" -Status "Unified Group Mailbox $i of $($UnifiedGroups.Count)" -PercentComplete (($i / $UnifiedGroups.Count) * 100)
    }
    $UnifiedMailboxStats = Get-MailboxStatistics -Identity $UnifiedGroup.PrimarySMTPAddress | Select-Object LastLogonTime, DisplayName, @{Name = "TotalItemSizeMB"; Expression = { [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }, ItemCount, DeletedItemCount, @{Name = "TotalDeletedItemSizeMB"; Expression = { [math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }
    $UnifiedGroupMailboxesArray = $UnifiedGroupMailboxesArray + [PSCustomObject]@{
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
    }
    $i++
}
Start-Sleep 5

    Write-Host "Writing Unified Group Mailbox data"
    $UnifiedGroupMailboxesArray | Export-CSV .\csvfiles\GroupMailbox.CSV -NoTypeInformation  


Write-Host "Unified Group Mailbox Data Exported"

Pause