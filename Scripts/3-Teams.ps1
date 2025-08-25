$Script = "Teams"
$Version = "v2.0.0"
Write-Host "Running" $Script $Version -ForegroundColor Green


Connect-MicrosoftTeams

$AllTeams = @()
$AllTeams = Get-Team | Sort-Object Displayname


$TeamArray = @()
$i = 1
Foreach ($Team in $AllTeams) {  
    If ($AllTeams.Count -gt "1") {
        Write-Progress  -Activity "Processing Microsoft Teams" -Status "Microsoft Team $i of $($AllTeams.Count)" -PercentComplete (($i / $AllTeams.Count) * 100) 
    }
    $Channels = Get-TeamChannel -GroupId $Team.GroupId
    $ChannelsCount = $Channels | Measure-Object
    $TeamArray = $TeamArray + [PSCustomObject]@{
        GroupId             = $Team.GroupId ; 
        DisplayName         = $Team.DisplayName ; 
        Description         = $Team.Description ; 
        Visibility          = $Team.Visibility ; 
        MailNickName        = $Team.MailNickName ; 
        Classification      = $Team.Classification ; 
        Archived            = $Team.Archived ; 
        Channels            = $ChannelsCount.Count ; 
    }
    $i++
}
Start-Sleep 5

    Write-Host "Writing Teams data"    
    $TeamArray | Export-CSV .\csvfiles\Teams.CSV -NoTypeInformation  


Write-host "Please wait while channel data is processed" -ForegroundColor Green
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
            }
        }
    }
    $i++
}
Start-Sleep 5
Write-Host "Writing Channel Data"    
$TeamChannelArray | Export-CSV .\csvfiles\TeamChannels.csv -NoTypeInformation  
Write-Progress  -ID 1 -Activity "Processing Microsoft Teams Channels" -Completed
$P++

Write-Host "Teams data exported"

Pause