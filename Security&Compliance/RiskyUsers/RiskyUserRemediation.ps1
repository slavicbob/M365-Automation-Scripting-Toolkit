<#
.SYNOPSIS
    RiskyUserRemediation.ps1 — Targeted remediation actions for a risky or compromised user account.
 
.DESCRIPTION
    This PowerShell script performs security remediation on a user flagged as risky in
    Microsoft Entra ID Identity Protection or via a manual security review. Unlike a full
    lockdown, this script takes a non-destructive remediation approach:
 
    - Generating a randomized password for potential on-premises reset use.
    - Installing and importing required PowerShell modules (Microsoft.Graph, ActiveDirectory).
    - Authenticating to Microsoft Graph via certificate-based app authentication.
    - Resolving the target user by UPN or proxy address (fallback lookup).
    - Forcing the user to change their on-premises AD password at next sign-in.
    - Revoking all active Entra ID sign-in sessions via the Graph API.
    - Outputting a structured JSON result object summarizing every action taken.
 
    Note: An optional cloud-side forced password change block (Update-MgUser) is included
    as a commented-out step and can be enabled as needed.
 
.PARAMETER UserPrincipalName
    The UPN or primary SMTP address of the risky user to remediate.
    Can be passed as a parameter or set directly in the CONFIGURATION block.
 
.OUTPUTS
    A JSON result object containing per-step status, error messages, and overall script status.
 
.EXAMPLE
    .\RiskyUserRemediation.ps1
 
    # Or uncomment the param block and run with a parameter:
    # .\RiskyUserRemediation.ps1 -UserPrincipalName "user@<YOUR_DOMAIN>"
 
.NOTES
    Required modules:
        - Microsoft.Graph (Authentication, Users, Identity.SignIns)
        - ActiveDirectory (RSAT)
 
    Required Entra ID app registration permissions:
        - User.ReadWrite.All
        - Directory.ReadWrite.All
 
    The $cred variable must be pre-populated with AD credentials that have write access
    to the target domain before running this script.
 
.ChangeLog
    V1.00 - Initial version
#>
 
# param(
#     [Parameter(Mandatory = $true)]
#     [string]$UserPrincipalName
# )
 
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
 
# ─────────────────────────────────────────────
# CONFIGURATION — replace all placeholder values
# ─────────────────────────────────────────────
 
# $UserPrincipalName = "target.user@<YOUR_DOMAIN>"
 
# Entra ID app registration (certificate-based auth)
$ClientId   = "<YOUR_CLIENT_ID>"
$TenantId   = "<YOUR_TENANT_ID>"
$Thumbprint = "<YOUR_CERT_THUMBPRINT>"
 
# Primary on-premises AD domain
$PrimaryADDomain = "<YOUR_PRIMARY_AD_DOMAIN>"
 
# AD credential with write access to the target domain (must be pre-populated)
# $cred = Get-Credential
 
# ─────────────────────────────────────────────
# GENERATE RANDOM PASSWORD
# ─────────────────────────────────────────────
 
$password  = -join ('abcdefghkmnpqrstuvwxyz'.ToCharArray() | Get-Random -Count 3)
$password += -join ('ABCDEFGHKLMNPRSTUVWXYZ'.ToCharArray() | Get-Random -Count 3)
$password += -join ('#*!@$%^&'.ToCharArray()               | Get-Random -Count 2)
$password += -join ('123456789'.ToCharArray()               | Get-Random -Count 2)
$password  = -join ($password.ToCharArray() | Sort-Object { Get-Random })
 
# ─────────────────────────────────────────────
# INITIALIZE RESULT OBJECT
# ─────────────────────────────────────────────
 
$Result = @{
    ScriptName   = "RiskyUserRemediation.ps1"
    User         = $UserPrincipalName
    TimeStamp    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    Status       = "InProgress"
    Actions      = @()
    Errors       = @()
    ModuleStatus = @()
}
 
try {
 
    # ─────────────────────────────────────────────
    # MODULE INSTALL / IMPORT
    # ─────────────────────────────────────────────
 
    foreach ($module in @("Microsoft.Graph", "ActiveDirectory")) {
        try {
            if (-not (Get-Module -ListAvailable -Name $module)) {
                Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                $Result.ModuleStatus += "$module installed successfully."
            }
            else {
                $Result.ModuleStatus += "$module already installed."
            }
            $Result.ModuleStatus += "$module imported successfully."
        }
        catch {
            $Result.ModuleStatus += "$module error: $($_.Exception.Message)"
            throw
        }
    }
 
    Import-Module Microsoft.Graph.Authentication   -ErrorAction Stop
    Import-Module Microsoft.Graph.Users            -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.SignIns -ErrorAction Stop
    Import-Module ActiveDirectory                  -ErrorAction Stop
 
    # ─────────────────────────────────────────────
    # AUTHENTICATE TO MICROSOFT GRAPH
    # ─────────────────────────────────────────────
 
    try {
        Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $Thumbprint -ErrorAction Stop -NoWelcome
        $Result.Actions += @{ Step = "Authentication"; Status = "Success"; Error = $null }
    }
    catch {
        $Result.Actions += @{ Step = "Authentication"; Status = "Failed"; Error = $_.Exception.Message }
        $Result.Errors  += $_.Exception.Message
        throw
    }
 
    # ─────────────────────────────────────────────
    # RESOLVE USER ACCOUNT
    # ─────────────────────────────────────────────
 
    try {
        $UserObj = Get-MgUser -UserId $UserPrincipalName -Property OnPremisesSecurityIdentifier, Id, UserPrincipalName |
                   Select-Object OnPremisesSecurityIdentifier, Id, UserPrincipalName -ErrorAction Stop
        $upn = $UserObj.UserPrincipalName
 
        if (-not $UserObj) {
            $UserObj = Get-MgUser -Filter "proxyAddresses/any(c:c eq 'smtp:$UserPrincipalName')" `
                           -Property OnPremisesSecurityIdentifier, Id, UserPrincipalName |
                       Select-Object OnPremisesSecurityIdentifier, Id, UserPrincipalName -ErrorAction Stop
            $upn = $UserObj.UserPrincipalName
 
            if (-not $UserObj) {
                $errorMsg = "User '$UserPrincipalName' not found by UserId or proxyAddresses filter."
                $Result.Actions += @{ Step = "Checking Account Existence"; Status = "Failed"; Error = $errorMsg }
                $Result.Errors  += $errorMsg
                throw
            }
        }
    }
    catch {
        $Result.Actions += @{ Step = "Checking Account Existence"; Status = "Failed"; Error = $_.Exception.Message }
        $Result.Errors  += $_.Exception.Message
        throw
    }
 
    # ─────────────────────────────────────────────
    # STEP 1 — FORCE ON-PREM PASSWORD CHANGE AT NEXT SIGN-IN
    # ─────────────────────────────────────────────
 
    try {
        $userobjpp = Get-ADUser -Identity $UserObj.OnPremisesSecurityIdentifier -Server $PrimaryADDomain -Credential $cred
        Set-ADUser -Identity $userobjpp.SID -ChangePasswordAtLogon $true
        $Result.Actions += @{ Step = "Force onprem change password next sign in"; Status = "Success"; Error = $null }
    }
    catch {
        $Result.Actions += @{ Step = "Force onprem change password next sign in"; Status = "Failed"; Error = $_.Exception.Message }
        $Result.Errors  += $_.Exception.Message
        throw
    }
 
    # ─────────────────────────────────────────────
    # OPTIONAL — FORCE CLOUD PASSWORD CHANGE (uncomment to enable)
    # ─────────────────────────────────────────────
 
    <#
    try {
        Update-MgUser -UserId $upn -PasswordProfile @{
            ForceChangePasswordNextSignIn = $true
            Password = $password
        }
        $Result.Actions += @{ Step = "Force change password next sign in"; Status = "Success"; Error = $null }
    }
    catch {
        $Result.Actions += @{ Step = "Force change password next sign in"; Status = "Failed"; Error = $_.Exception.Message }
        $Result.Errors  += $_.Exception.Message
        throw
    }
    #>
 
    # ─────────────────────────────────────────────
    # STEP 2 — REVOKE ALL ACTIVE SIGN-IN SESSIONS
    # ─────────────────────────────────────────────
 
    try {
        $uri = "https://graph.microsoft.com/v1.0/users/$upn/revokeSignInSessions"
        Invoke-MgGraphRequest -Method POST -Uri $uri -ErrorAction Stop | Out-Null
        $Result.Actions += @{ Step = "Revoke Sign-In Sessions"; Status = "Success"; Error = $null }
    }
    catch {
        $Result.Actions += @{ Step = "Revoke Sign-In Sessions"; Status = "Failed"; Error = $_.Exception.Message }
        $Result.Errors  += $_.Exception.Message
        throw
    }
 
}
catch {
    $Result.Actions += @{ Step = "Script Execution"; Status = "Failed"; Error = $_.Exception.Message }
    $Result.Errors  += $_.Exception.Message
    $Result.Status   = "Failed"
}
finally {
    if ($Result.Status -eq "InProgress") { $Result.Status = "CompletedSuccessfully" }
    Disconnect-MgGraph | Out-Null
    $Result | ConvertTo-Json -Depth 5
}