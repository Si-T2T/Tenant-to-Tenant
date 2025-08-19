<#




    Install-Module Microsoft.Graph -scope AllUsers -Force
    Install-Module ExchangeOnlineManagement -Force
    Install-Module MicrosoftTeams -Force
    Install-Module ImportExcel -Force
    Install-Module Microsoft.Online.SharePoint.PowerShell -Force



#>
$Script = "BootStapper"
$Version = "v2.0.0"
Write-Host "Running" $Script $Version -ForegroundColor Green


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

#region Functions

function CheckForPowerShellModule([string]$ModuleName)
{
#  Write-Host "Checking for $ModuleName module..."
  if (Get-InstalledModule -Name $ModuleName)
  {
    # Module already installed 
    Write-Host "$ModuleName installed - Continuing" -ForegroundColor Green
  }
  else
  {
    # Module not installed
    Write-Host "$ModuleName not found. Please install $ModuleName module and rerun script" -ForegroundColor Red
    Break
  }
}
#endregion Functions

#region Check for Modules
CheckForPowerShellModule("Microsoft.Graph")
CheckForPowerShellModule("ExchangeOnlineManagement")
CheckForPowerShellModule("ImportExcel")
CheckForPowerShellModule("Microsoft.Online.SharePoint.PowerShell")
CheckForPowerShellModule("MicrosoftTeams")
#CheckForPowerShellModule("Microsoft.PowerApps.Administration.PowerShell")

#endregion Check Modules

#region Variables and Template Check

$ScriptVersion = "2.0.0"
$TemplateVersion = "2.0.0"
$Template = ""
$Template = "Template_" + $TemplateVersion + ".xlsx"

# Check if the Template exists
if (-not (Test-Path -Path $Template)) {
    Write-Host "The file" $Template "does not exist. Ensure the template is in the working directory and restert the script" -ForegroundColor Red
    Break
}
Write-Host "The file" $Template "exists. Continuing with the script..." -ForegroundColor Green

if (-not (Test-Path -Path .\CSVFiles)) {
    Write-Host "The Folder" $Template "does not exist. Creating folder" -ForegroundColor Yellow
    New-Item -Name "CSVFiles" -ItemType "directory"
}
Else {
        Write-Host "Folder CSVFiles exists. Continuing....." -ForegroundColor Green
}


$Consultant = Read-Host "Please Enter Your Name and press Enter"
$Consultant | Out-File .\CSVFiles\Consultant.txt

#endregion Variables and Template Check

#region Graph Stuff

#endregion Graph Stuff

Start-Process pwsh -ArgumentList "-NoProfile -File .\MGGraph.ps1"
Sleep 30
Start-Process pwsh -ArgumentList "-NoProfile -File .\EXOBootStrapper.ps1"
Sleep 30
Start-Process pwsh -ArgumentList "-NoProfile -File .\Teams.ps1"
Sleep 30
Start-Process pwsh -ArgumentList "-NoProfile -File .\SPOnline.ps1"

While (!(Test-Path .\csvfiles\AllMGUsers.csv -ErrorAction SilentlyContinue))
{
  # endless loop, when the file will be there, it will continue
}
While (!(Test-Path .\csvfiles\AccountMailbox.csv -ErrorAction SilentlyContinue))
{
  # endless loop, when the file will be there, it will continue
}
While (!(Test-Path .\csvfiles\GroupMailbox.csv -ErrorAction SilentlyContinue))
{
  # endless loop, when the file will be there, it will continue
}
While (!(Test-Path .\csvfiles\UnifiedGroups.csv -ErrorAction SilentlyContinue))
{
  # endless loop, when the file will be there, it will continue
}
While (!(Test-Path .\csvfiles\Teams.csv -ErrorAction SilentlyContinue))
{
  # endless loop, when the file will be there, it will continue
}
While (!(Test-Path .\csvfiles\TeamChannels.csv -ErrorAction SilentlyContinue))
{
  # endless loop, when the file will be there, it will continue
}
While (!(Test-Path .\csvfiles\OneDrives.csv -ErrorAction SilentlyContinue))
{
  # endless loop, when the file will be there, it will continue
}
While (!(Test-Path .\csvfiles\SPOnline.csv -ErrorAction SilentlyContinue))
{
  # endless loop, when the file will be there, it will continue
}





Write-Host "Script Complete, good knob" $Consultant

Pause