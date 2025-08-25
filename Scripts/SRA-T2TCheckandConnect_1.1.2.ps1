<#
        SRA-T2TCheckandConnect_1.1.2.PS1
        v1.0.1
        09/12/2024
        simonan@softcat.com        
        
        .SYNOPSIS
        Connects to Microsoft 365 so the T2T Discovery script can be run

        .DESCRIPTION
        This PowerShell script connects to Microsoft 365:
            Graph
            Exchange Online Management
            SharePoint Online
            Microsoft Teams
            Power BI Service

        .PARAMETER 
        
        .PARAMETER 
        
        .INPUTS
        
        .OUTPUTS
        
        .EXAMPLE


        .LINK
        https://www.softcat.com

#>

Write-Host "" 
Write-Host "Disclaimer:" -ForegroundColor Red
Write-Host "
This PowerShell script is provided as is without any warranty of any kind, either express or implied, including but not 
limited to the implied warranties of merchantability and fitness for a particular purpose.
The entire risk as to the quality and performance of the script is with you. Should the script prove defective, you assume 
the cost of all necessary servicing, repair, or correction.
In no event shall the author or contributors be liable for any damages whatsoever (including, without limitation, damages 
for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of 
the use of or inability to use this script, even if the author has been advised of the possibility of such damages.
Use this script at your own risk." 
Write-Host "" 
    
$continue = Read-Host "Do you want to continue? (Y/N)"
if ($continue -ne "Y") {
    Write-Host "Script execution has been cancelled." -ForegroundColor Red
    Break
}
Write-Host ""   
    
# Check if connected to Microsoft Graph
function Check-GraphConnection {

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
    $graphConnection = @()    
    $graphConnection = Get-mgcontext  -ErrorAction SilentlyContinue
    if ($graphConnection -eq $Null) {
        Write-Output "Not connected to Microsoft Graph"
        Write-Host "Connecting to Graph PowerShell"
        Connect-MGGraph @Scope
    
        $Org = Get-MgOrganization
        Write-Host "" 
    
        $continue = Read-Host "You have connected Microsoft Graph to" $Org.DisplayName", Is this correct (Y/N)? If not, please disconnect ALL SESSIONS from Graph and reconnect"
        Write-Host "" 
        if ($continue -ne "Y") {
            Write-Host "Script execution has been cancelled." -ForegroundColor Red
            Break
        }
    
    }
    else {
        Write-Output "Connected to Microsoft Graph" 
        $Org = Get-MgOrganization
    Write-Host "" 
        $continue = Read-Host "You have connected Microsoft Graph to" $Org.DisplayName", Is this correct (Y/N)? If not, please disconnect ALL SESSIONS from Graph and reconnect (Disconnect-Graph, Disconnect-MGGraph)"
        if ($continue -ne "Y") {
            Write-Host "Script execution has been cancelled." -ForegroundColor Red
            Break
        }
    }
}
    
# Check if connected to Exchange Online
function Check-ExchangeConnection {
    try {
        $exchangeConnection = Get-AcceptedDomain 
        if ($exchangeConnection) {
            Write-Output "Connected to Exchange Online"
        }
    }
    catch {
        Write-Output "Not connected to Exchange Online"
        Write-Host "Connecting to Exchange Online Management"
        Connect-ExchangeOnline
    
    }
}
    
# Check if connected to SharePoint
function Check-SharePointConnection {
    try {
        $sharePointConnection = Get-SPOSite
        if ($sharePointConnection) {
            Write-Output "Connected to SharePoint"
        }
    }
    catch {
        Write-Output "Not connected to SharePoint"
        Write-Host "Connecting to SharePoint Online"
        $OnMicrosoftDomain = Get-AcceptedDomain | Where-Object { $_.DomainName -like "*.onmicrosoft.com" } | Select-Object DomainName -ExpandProperty DomainName
        $OnMicrosoftPrefix = $OnMicrosoftDomain.split('.')[0] 
        $AdminURL = "https://" + $OnMicrosoftPrefix + "-admin.sharepoint.com"
        Connect-SPOService -URL $AdminURL
    }
}
    
# Check if connected to Teams
function Check-TeamsConnection {
    try {
        $teamsConnection = Get-CsOnlineUser -ResultSize 1
        if ($teamsConnection) {
            Write-Output "Connected to Teams"
        }
    }
    catch {
        Write-Output "Not connected to Teams"
        Write-Host "Connecting to Microsoft Teams"
        Connect-MicrosoftTeams
    }
}
    
# Check if connected to Power BI
function Check-PowerBIConnection {
 

        Write-Host "Connecting to the Power BI service can fail if the service has not been activated!"  -ForegroundColor Yellow 
        $continue = Read-Host "I can't tell if you are connected to the Power BI Service. Do you want me to try and connect anyway (Y/N)? "

        if ($continue -eq "Y") {
            Write-Host "Connecting to the Power BI Service "
                    Connect-PowerBIServiceAccount
        }

        if ($continue -ne "Y") {
            Write-Host "Not even gonna try then!" -ForegroundColor Red
            Break
        }
    
}
    
# Run all checks
Function Check-Connections {
    Check-GraphConnection
    Check-ExchangeConnection
    Check-SharePointConnection
    Check-TeamsConnection
    Check-PowerBIConnection
}
    
Check-Connections