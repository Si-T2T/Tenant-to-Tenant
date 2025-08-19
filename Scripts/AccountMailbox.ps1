$Script = "AccountMailbox"
$Version = "v2.0.0"
Write-Host "Running" $Script $Version -ForegroundColor Green

Connect-ExchangeOnline -DisableWAM

Write-Host "Getting Accepted Domains" -ForegroundColor Green

$AcceptedDomains = Get-AcceptedDomain | Sort-Object ascending

Write-Host
$OnMicrosoftDomain = $AcceptedDomains| Where-Object { $_.DomainName -like "*.onmicrosoft.com" } | Select-Object DomainName -ExpandProperty DomainName
        $OnMicrosoftPrefix = $OnMicrosoftDomain.split('.')[0] 
        $AdminURL = "https://" + $OnMicrosoftPrefix + "-admin.sharepoint.com"
$AdminURL | Out-File .\CSVFiles\AdminURL.txt

$CustomDomains = $AcceptedDomains| Where-Object { $_.DomainName -notlike "*.onmicrosoft.com" } 
$CustomDomains | Export-csv .\CSVFiles\CustomDomains.csv -NoTypeInformation

Write-Host "Getting Mailboxes" -ForegroundColor Green

$AllMailboxes = @()
$AllMailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.PrimarySMTPAddress -notLike "DiscoverySearchMailbox*" } | Sort-Object UserPrincipalname

$AllMailboxesArray = @()
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
    }
    $i++
}

Start-Sleep 5

    $AllMailboxesArray | Export-CSV .\csvfiles\AccountMailbox.csv -NoTypeInformation 


Write-Host "Mailbox data exported"


Pause