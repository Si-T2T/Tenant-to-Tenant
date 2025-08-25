#Check if profiles exist
$profiles = @{
    "CurrentUserCurrentHost" = $PROFILE.CurrentUserCurrentHost
    "CurrentUserAllHosts"    = $PROFILE.CurrentUserAllHosts
    "AllUsersCurrentHost"    = $PROFILE.AllUsersCurrentHost
    "AllUsersAllHosts"       = $PROFILE.AllUsersAllHosts
}

# Check existence and output results
foreach ($key in $profiles.Keys) {
    $path = $profiles[$key]
    if (Test-Path $path) {
        Write-Host "$key profile exists at: $path" -ForegroundColor Green
    } else {
        Write-Host "$key profile does NOT exist. Expected path: $path" -ForegroundColor Red
    }
}
