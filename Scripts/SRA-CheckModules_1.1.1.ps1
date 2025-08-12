#region CheckForPowerShellModule
function CheckForPowerShellModule([string]$ModuleName)
{
#  Write-Host "Checking for $ModuleName module..."
  if (Get-Module -ListAvailable -Name $ModuleName)
  {
    # Module already installed 
    Write-Host "$ModuleName installed - Continuing" -ForegroundColor Green
  }
  else
  {
    # Module not installed
    Write-Host "$ModuleName not found. Please install $ModuleName module and rerun script" -ForegroundColor Red
    #Exit
  }
}
#endregion CheckForPowerShellModule
#endregion Functions

#region Check for Modules
CheckForPowerShellModule("Microsoft.Graph")
CheckForPowerShellModule("ExchangeOnlineManagement")
CheckForPowerShellModule("ImportExcel")
CheckForPowerShellModule("Microsoft.Online.SharePoint.PowerShell")
CheckForPowerShellModule("MicrosoftTeams")
CheckForPowerShellModule("Microsoft.PowerApps.Administration.PowerShell")
CheckForPowerShellModule("MicrosoftPowerBIMgmt")


#endregion Check for Modules
