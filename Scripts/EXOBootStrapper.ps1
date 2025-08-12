$Script = "EXOBootStrapper"
$Version = "v2.0.0"
Write-Host "Running" $Script $Version -ForegroundColor Green

Start-Process pwsh -ArgumentList "-NoProfile -File .\AccountMailbox.ps1"
While (!(Test-Path .\csvfiles\AccountMailbox.csv -ErrorAction SilentlyContinue))
{
  # endless loop, when the file will be there, it will continue
}

Write-Host "Mailbox Data Export Complete" -ForegroundColor Green

Start-Process pwsh -ArgumentList "-NoProfile -File .\UnifiedGroups.ps1"

While (!(Test-Path .\csvfiles\GroupMailbox.csv -ErrorAction SilentlyContinue))
{
  # endless loop, when the file will be there, it will continue
}

Write-Host "Unified Group Data Export Complete" -ForegroundColor Green

<#

Start-Process pwsh -ArgumentList "-NoProfile -File .\Groups.ps1"

Write-Host "Group Data Export Complete" -ForegroundColor Green
#>


pause