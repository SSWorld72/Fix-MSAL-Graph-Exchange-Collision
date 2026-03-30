<#
.SYNOPSIS
    Ultimate Workaround for MSAL DLL Version Collisions & WAM Crashes when mixing 
    Microsoft.Graph and ExchangeOnlineManagement modules.

.DESCRIPTION
    This script provides a reliable, reproducible pattern for connecting to both 
    Exchange Online and Microsoft Graph in the same session, avoiding the notoriously
    common Application Builder exceptions or NullReferenceExceptions.

    Background:
    1. If you run both modules in PowerShell 5.1 (Windows PowerShell), their 
       underlying Microsoft.Identity.Client.dll versions clash in the single AppDomain,
       throwing "Method not found: WithBroker", "WithLogging", or "CouldNotAutoloadMatchingModule".
    2. If you migrate to PowerShell 7 (pwsh) to isolate the DLLs via AssemblyLoadContext, 
       ExchangeOnlineManagement V3.4+ has a fatal native bug. When you supply the 
       -Credential parameter silently (ROPC flow), Microsoft's Web Account Manager (WAM) 
       Broker code attempts to execute without a UI handle, triggering an immediate 
       and uncatchable "System.NullReferenceException: Object reference not set... 
       at RuntimeBroker..ctor".
    3. Furthermore, most modern Entra ID (Azure AD) tenants enforce MFA (AADSTS50076), 
       making legacy -Credential (ROPC without interactive browser) obsolete anyway.

    The Fix:
    - Run this in PowerShell 7 (pwsh) ONLY.
    - Set WAM environment variables to "false" BEFORE importing Exchange module.
    - Import and Connect Exchange FIRST, to avoid MSAL initialization conflicts.
    - DO NOT use the `-Credential` parameter in `Connect-ExchangeOnline`. 
      Use `-UserPrincipalName` to elegantly force the interactive system browser (which natively supports MFA/SSO) 
      and bypasses the broken PS7 WAM component entirely.
    - Import and Connect Microsoft Graph AFTER Exchange successfully connects.

.EXAMPLE
    pwsh.exe -File .\Fix-MSAL-Collision-Graph-Exchange.ps1
#>

# 1. Enforce PowerShell 7 (Core) as a prerequisite to isolate MSAL DLLs
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Warning "This script MUST be run in PowerShell 7 (pwsh.exe) to avoid AppDomain MSAL DLL collisions."
    Write-Warning "Please install PowerShell 7: winget install Microsoft.PowerShell"
    exit
}

Write-Host "`n============== Connecting Microsoft 365 Services ==============`n" -ForegroundColor Yellow

# 2. Prevent Graph from implicitly auto-loading its MSAL DLL before Exchange.
# If Microsoft.Graph.Authentication is already loaded in memory, it will break Exchange's initialization.
if (Get-Module -Name Microsoft.Graph* -ErrorAction SilentlyContinue) {
    Write-Warning "Graph modules are already loaded. Fresh session is highly recommended."
    Remove-Module Microsoft.Graph* -Force -ErrorAction SilentlyContinue 
}

# 3. Disable the bugged WAM Broker feature in MSAL securely (Requires PS7 strings, not booleans)
$env:EXO_UseWindowsBroker = "false"
$env:MSAL_FORCE_USERAGENT = "1"

Write-Host "`n [Exchange] Connecting...`n" -NoNewline -ForegroundColor Cyan

try {
    # 4. Ensure ExchangeOnlineManagement is installed
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host " ExchangeOnlineManagement module not found. Installing..." -ForegroundColor DarkGray
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }

    # 5. Import ExchangeOnlineManagement FIRST
    Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop

    # 5. Connect Exchange WITHOUT `-Credential` param but WITH `-UserPrincipalName`!
    # By providing the UPN dynamically, MSAL natively skips the buggy Windows WAM OS picker 
    # (which throws AADSTS1400011 Native App redirect URI bugs) and correctly jumps straight
    # to the standard, foolproof Web Browser login!
    
    $adminUPN = Read-Host "Please enter your Admin Email (UPN) to start Web login"
    Connect-ExchangeOnline -UserPrincipalName $adminUPN -ShowProgress:$false -ShowBanner:$false -ErrorAction Stop
    
    Write-Host " Connected successfully." -ForegroundColor Green
}
catch {
    Write-Host "`n [Error] Exchange connection failed: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

Write-Host "`n [Graph]    Connecting..." -NoNewline -ForegroundColor Cyan

try {
    # 6. Ensure Microsoft.Graph.Authentication is installed
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host "`n Microsoft.Graph module not found. Installing..." -ForegroundColor DarkGray
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }

    # 7. Import Graph ONLY AFTER Exchange has connected securely
    Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop
    
    # 7. Connect Graph
    Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All" -NoWelcome -ErrorAction Stop
    
    Write-Host "`n Connected successfully." -ForegroundColor Green
}
catch {
    Write-Host "`n [Error] Graph connection failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n==============================================================`n" -ForegroundColor Yellow

# Test actual cmdlets to prove stability
try {
    Write-Host " [Test 1] Testing Exchange Cmdlet: Get-Mailbox..." -ForegroundColor Gray
    $mbx = Get-Mailbox -ResultSize 1 -WarningAction SilentlyContinue -ErrorAction Stop
    $maskedMbx = $mbx.PrimarySmtpAddress -replace "(?<=^.{1}).*(?=@)", "***" -replace "(?<=@)[^.]+", "***"
    Write-Host "  > Found mailbox: $maskedMbx`n" -ForegroundColor DarkCyan
    
    Write-Host " [Test 2] Testing Graph Cmdlet: Get-MgUser..." -ForegroundColor Gray
    $user = Get-MgUser -Top 1 -WarningAction SilentlyContinue -ErrorAction Stop
    $maskedUser = $user.UserPrincipalName -replace "(?<=^.{1}).*(?=@)", "***" -replace "(?<=@)[^.]+", "***"
    Write-Host "  > Found user:    $maskedUser`n" -ForegroundColor DarkCyan
    
    Write-Host " ==============================================================" -ForegroundColor Yellow
    Write-Host " 🎉 Congratulations! Both modules are running perfectly together." -ForegroundColor Green
}
catch {
    Write-Host "`n [Error] Cmdlet execution failed: $($_.Exception.Message)" -ForegroundColor Red
}
