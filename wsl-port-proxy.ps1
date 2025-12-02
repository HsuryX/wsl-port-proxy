<#
.SYNOPSIS
    Automates WSL Port Proxy and firewall rules setup, ensuring cleanup of existing and deprecated rules.
.DESCRIPTION
    - Creates a WSL Port Proxy configuration script.
    - Registers a scheduled task to run the script at system startup using a specified user account.
    - Cleans up and configures new inbound firewall rules for defined ports.
.NOTES
    Requires the script to be run with elevated permissions (Administrator).
    You will be prompted to enter the password for the scheduled task user.
#>
#Requires -RunAsAdministrator

# =================================================================================
# 1. Configuration
# =================================================================================

$WslPorts        = @(22, 80, 443, 3000)
$ScriptDirectory = "C:\Scripts"
$ScriptFileName  = "wsl-port-proxy.ps1"
$ScriptPath      = Join-Path -Path $ScriptDirectory -ChildPath $ScriptFileName
$TaskName        = "WSL Port Proxy"

# =================================================================================
# 2. Create directory and Port Proxy configuration script
# =================================================================================

Write-Host "## 1. Preparing script directory and writing configuration file..."

if (-not (Test-Path $ScriptDirectory)) {
    New-Item -ItemType Directory -Path $ScriptDirectory | Out-Null
    Write-Host "Directory created: $ScriptDirectory"
} else {
    Write-Host "Directory already exists: $ScriptDirectory"
}

$PortsText = $WslPorts -join ", "

$PortProxyScriptContent = @'
<#
.SYNOPSIS
    Configures WSL Port Proxy rules via netsh.
#>
$ErrorActionPreference = 'Stop'
$Ports = @({PORTS})

try {
    $WslIp = wsl.exe hostname -I 2>$null
    if (-not $WslIp) {
        Write-Warning "Could not retrieve WSL IP address. Skipping port forwarding."
        exit 1
    }
    $WslIp = ($WslIp.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries) | Select-Object -First 1).Trim()
}
catch {
    Write-Warning "Failed to run wsl.exe. Skipping port forwarding. Error: $($_.Exception.Message)"
    exit 1
}

Write-Output "Current WSL IP Address: $WslIp"

foreach ($Port in $Ports) {
    netsh interface portproxy delete v4tov4 listenport=$Port listenaddress=0.0.0.0 2>$null | Out-Null
    netsh interface portproxy add v4tov4 listenport=$Port listenaddress=0.0.0.0 connectport=$Port connectaddress=$WslIp
    
    if ($LASTEXITCODE -eq 0) {
        Write-Output "Port $Port successfully forwarded to $WslIp"
    } else {
        Write-Warning "Port $Port forwarding failed. netsh exit code: $LASTEXITCODE"
    }
}
'@

$PortProxyScriptContent = $PortProxyScriptContent -replace '\{PORTS\}', $PortsText

Set-Content -Path $ScriptPath -Value $PortProxyScriptContent -Force -Encoding UTF8
Write-Host "WSL Port Proxy setup script created or updated at $ScriptPath"

# =================================================================================
# 3. Register Scheduled Task
# =================================================================================

Write-Host "## 2. Registering system startup scheduled task..."

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null

$CurrentUser = "$env:UserDomain\$env:UserName"
Write-Host "Please enter the password for the scheduled task user (defaulting to current user: $CurrentUser):"

try {
    $Credential = Get-Credential -UserName $CurrentUser -Message "Input the password for user to run the task"
}
catch {
    Write-Error "Failed to retrieve credentials. Aborting task registration."
    exit 1
}

$Action    = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""
$Trigger   = New-ScheduledTaskTrigger -AtStartup
$Principal = New-ScheduledTaskPrincipal -UserId $Credential.UserName -RunLevel Highest -LogonType Password

try {
    Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -User $Credential.UserName -Password $Credential.GetNetworkCredential().Password -RunLevel Highest -Force
    Write-Host "Scheduled task '$TaskName' registered to run at system startup using user: $($Credential.UserName)."

    Start-ScheduledTask -TaskName $TaskName
    Write-Host "Scheduled task '$TaskName' triggered."
}
catch {
    Write-Error "Failed to register scheduled task: $($_.Exception.Message)"
}

# =================================================================================
# 4. Cleanup Specific Firewall Rules
# =================================================================================

Write-Host "## 3. Cleaning up existing specific Firewall Rules..."

$RulesToDelete = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "WSL-Inbound-TCP-*" }
$DeletedCount  = 0

foreach ($Rule in $RulesToDelete) {
    Remove-NetFirewallRule -Name $Rule.Name -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "   Deleted Existing Rule: $($Rule.DisplayName)"
    $DeletedCount++
}

if ($DeletedCount -eq 0) {
    Write-Host "Cleanup complete: No existing specific rules required removal."
} else {
    Write-Host "Cleanup complete: Total $DeletedCount specific rules removed."
}

# =================================================================================
# 5. Firewall rules configuration
# =================================================================================

Write-Host "## 4. Configuring New Windows Firewall rules..."

foreach ($Port in $WslPorts) {
    $FirewallRuleName = "WSL-Inbound-TCP-$Port"
    
    if (-not (Get-NetFirewallRule -DisplayName $FirewallRuleName -ErrorAction SilentlyContinue)) {
        New-NetFirewallRule -DisplayName $FirewallRuleName -Direction Inbound -Protocol TCP -LocalPort $Port -Action Allow -Profile Any -Name $FirewallRuleName
        Write-Host "Firewall rule added: $FirewallRuleName (Port $Port)"
    } else {
        Write-Host "Warning: Firewall rule $FirewallRuleName unexpectedly exists. Skipping creation."
    }
}

# =================================================================================
# 6. Completion
# =================================================================================

Write-Host "WSL Port Proxy Setup Complete."
