<#
.SYNOPSIS
    UserLockdown.ps1 — Immediate full account lockdown for a target M365/AD user(Inspired and built in a hybrid environment).
 
.DESCRIPTION
    This PowerShell script performs a complete security lockdown of a user account across
    both cloud (Entra ID / Microsoft 365) and on-premises Active Directory environments.
    It is designed for rapid response scenarios such as employee offboarding, security
    incidents, or account compromise.
 
    The script automates the following actions in sequence:
    - Generating a cryptographically randomized replacement password.
    - Installing and importing required PowerShell modules (Microsoft.Graph, ActiveDirectory).
    - Authenticating to Microsoft Graph via certificate-based app authentication.
    - Resolving the target user by UPN or proxy address (fallback lookup).
    - Disabling the user's Entra ID (cloud) account.
    - Disabling the user's primary on-premises AD account and resetting its password.
    - Enumerating and disabling the user's accounts across additional AD forests/domains.
    - Revoking all active sign-in sessions via the Graph API.
    - Removing all registered MFA methods (Email, Phone, FIDO2, Microsoft Authenticator,
      Software OATH, Windows Hello for Business, Temporary Access Pass).
    - Outputting a structured, filtered JSON result object summarizing every action taken.
 
.PARAMETER UserPrincipalName
    The UPN or primary SMTP address of the user to lock down.
    Can be passed as a parameter or set directly in the CONFIGURATION block.
 
.OUTPUTS
    A filtered JSON result object containing per-step status, error messages, MFA revocation
    summary (counts + method lists), and on-prem domain results.
 
.EXAMPLE
    # Run interactively with a hardcoded UPN:
    .\UserLockdown.ps1
 
    # Or uncomment the param block and run with a parameter:
    # .\UserLockdown.ps1 -UserPrincipalName "user@<YOUR_DOMAIN>"
 
.NOTES
    Required modules:
        - Microsoft.Graph (Authentication, Users, Identity.SignIns)
        - ActiveDirectory (RSAT)
 
    Required Entra ID app registration permissions:
        - User.ReadWrite.All
        - UserAuthenticationMethod.ReadWrite.All
        - Directory.ReadWrite.All
 
    The service principal must authenticate via a certificate (thumbprint-based).
    The $cred variable must be pre-populated with credentials that have AD write access
    across all target domains/forests before running this script.
 
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
 
# Target user
$UserPrincipalName = "target.user@<YOUR_DOMAIN>"
 
# Entra ID app registration (certificate-based auth)
$ClientId    = "<YOUR_CLIENT_ID>"
$TenantId    = "<YOUR_TENANT_ID>"
$Thumbprint  = "<YOUR_CERT_THUMBPRINT>"
 
# Primary on-premises AD domain
$PrimaryADDomain = "<YOUR_PRIMARY_AD_DOMAIN>"
 
# Additional forests to search for linked accounts
$AdditionalForests = @("<YOUR_FOREST_1>", "<YOUR_FOREST_2>")
 
# Custom AD attribute used as an alternate UPN identifier in your environment
$CustomUPNAttribute = "<YOUR_CUSTOM_AD_UPN_ATTRIBUTE>"
 
# AD credential with write access across all target domains (must be pre-populated)
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
    ScriptName          = "UserLockdown.ps1"
    User                = $UserPrincipalName
    TimeStamp           = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    Status              = "InProgress"
    Actions             = @()
    Errors              = @()
    ModuleStatus        = @()
    NewPassword         = $password
    OnPremDomainResults = @{
        Success = @()
        Failed  = @()
    }
    MFAResults = @{
        Success = @()
        Failed  = @()
        Skipped = @()
    }
}
 
# ─────────────────────────────────────────────
# MODULE INSTALL / IMPORT
# ─────────────────────────────────────────────
 
try {
 
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
 
    Import-Module Microsoft.Graph.Authentication      -ErrorAction Stop
    Import-Module Microsoft.Graph.Users               -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.SignIns    -ErrorAction Stop
    Import-Module ActiveDirectory                     -ErrorAction Stop
 
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
    # STEP 1 — DISABLE CLOUD (ENTRA ID) ACCOUNT
    # ─────────────────────────────────────────────
 
    try {
        $uri  = "https://graph.microsoft.com/v1.0/users/$upn"
        $body = @{ accountEnabled = $false } | ConvertTo-Json
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $body -ErrorAction Stop | Out-Null
        $Result.Actions += @{ Step = "Disable Cloud Account"; Status = "Success"; Error = $null }
    }
    catch {
        $Result.Actions += @{ Step = "Disable Cloud Account"; Status = "Failed"; Error = $_.Exception.Message }
        $Result.Errors  += $_.Exception.Message
        throw
    }
 
    # ─────────────────────────────────────────────
    # STEP 1.1 — DISABLE ON-PREM ACCOUNTS & RESET PASSWORD
    # ─────────────────────────────────────────────
 
    try {
        $userobjpp = Get-ADUser -Identity $UserObj.OnPremisesSecurityIdentifier -Server $PrimaryADDomain -Credential $cred
 
        # Disable primary on-prem account
        try {
            Disable-ADAccount -Identity $userobjpp.SID -Server $PrimaryADDomain -Credential $cred
            $Result.Actions += @{ Step = "Disable Primary On-prem Account"; Status = "Success" }
        }
        catch {
            $Result.Actions += @{ Step = "Disable Primary On-prem Account"; Status = "Failed"; Error = $_.Exception.Message }
            $Result.Errors  += $_.Exception.Message
            throw
        }
 
        # Reset on-prem password
        try {
            Set-ADAccountPassword -Identity $userobjpp -NewPassword (ConvertTo-SecureString $password -AsPlainText -Force) -Reset -Confirm:$false -Credential $cred
            $Result.Actions += @{ Step = "Reset On-prem Password"; Status = "Success" }
        }
        catch {
            $Result.Actions += @{ Step = "Reset On-prem Password"; Status = "Failed"; Error = $_.Exception.Message }
            $Result.Errors  += $_.Exception.Message
            throw
        }
 
        # Enumerate domains across additional forests and disable linked accounts
        $domains = foreach ($forest in $AdditionalForests) {
            (Get-ADForest -Server $forest -Credential $cred).domains
        }
 
        foreach ($domain in $domains) {
            try {
                $aduser = Get-ADUser -Filter {
                    (EmailAddress -eq $UserPrincipalName) -or
                    ($CustomUPNAttribute -eq $UserPrincipalName) -or
                    (userPrincipalName -eq $UserPrincipalName)
                } -Server $domain -Credential $cred
 
                if ($aduser) {
                    try {
                        Disable-ADAccount -Identity $aduser.SID -Server $domain -Credential $cred
                        $Result.Actions += @{ Step = "Disable Additional Domain Account ($domain)"; Status = "Success" }
                        $Result.OnPremDomainResults.Success += $domain
                        break
                    }
                    catch {
                        $Result.Actions += @{ Step = "Disable Additional Domain Account ($domain)"; Status = "Failed"; Error = $_.Exception.Message }
                        $Result.Errors  += $_.Exception.Message
                        $Result.OnPremDomainResults.Failed += $domain
                        throw
                    }
                }
            }
            catch {
                $Result.Actions += @{ Step = "Lookup Additional Domain Account ($domain)"; Status = "Failed"; Error = $_.Exception.Message }
                $Result.Errors  += $_.Exception.Message
                $Result.OnPremDomainResults.Failed += $domain
            }
        }
    }
    catch {
        $Result.Actions += @{ Step = "Disable On-prem Account(s)"; Status = "Failed"; Error = $_.Exception.Message }
        $Result.Errors  += $_.Exception.Message
        throw
    }
 
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
 
    # ─────────────────────────────────────────────
    # STEP 3 — REMOVE ALL MFA METHODS
    # ─────────────────────────────────────────────
 
    try {
        $UserId       = $UserObj.Id
        $revokedCount = 0
        $skippedCount = 0
        $failedCount  = 0
 
        $Methods = Get-MgUserAuthenticationMethod -UserId $UserId -ErrorAction Stop
 
        if (-not $Methods) {
            $Result.Actions += @{ Step = "Revoke MFA Methods"; Status = "Skipped"; Error = "No MFA methods found" }
        }
        else {
            foreach ($method in $Methods) {
                $methodType = $method.AdditionalProperties['@odata.type']
                $methodId   = $method.Id
 
                switch ($methodType) {
 
                    "#microsoft.graph.emailAuthenticationMethod" {
                        try {
                            $null = Remove-MgUserAuthenticationEmailMethod -UserId $UserId -EmailAuthenticationMethodId $methodId -ErrorAction Stop
                            $Result.Actions += @{ Step = "Revoke MFA Method (Email)"; Status = "Success"; Error = $null }
                            $Result.MFAResults.Success += "Email"; $revokedCount++
                        }
                        catch {
                            $Result.Actions += @{ Step = "Revoke MFA Method (Email)"; Status = "Failed"; Error = $_.Exception.Message }
                            $Result.Errors  += $_.Exception.Message
                            $Result.MFAResults.Failed += "Email"; $failedCount++
                        }
                    }
 
                    "#microsoft.graph.fido2AuthenticationMethod" {
                        try {
                            $null = Remove-MgUserAuthenticationFido2Method -UserId $UserId -Fido2AuthenticationMethodId $methodId -ErrorAction Stop
                            $Result.Actions += @{ Step = "Revoke MFA Method (FIDO2)"; Status = "Success"; Error = $null }
                            $Result.MFAResults.Success += "FIDO2"; $revokedCount++
                        }
                        catch {
                            $Result.Actions += @{ Step = "Revoke MFA Method (FIDO2)"; Status = "Failed"; Error = $_.Exception.Message }
                            $Result.Errors  += $_.Exception.Message
                            $Result.MFAResults.Failed += "FIDO2"; $failedCount++
                        }
                    }
 
                    "#microsoft.graph.phoneAuthenticationMethod" {
                        try {
                            $null = Remove-MgUserAuthenticationPhoneMethod -UserId $UserId -PhoneAuthenticationMethodId $methodId -ErrorAction Stop
                            $Result.Actions += @{ Step = "Revoke MFA Method (Phone)"; Status = "Success"; Error = $null }
                            $Result.MFAResults.Success += "Phone"; $revokedCount++
                        }
                        catch {
                            $Result.Actions += @{ Step = "Revoke MFA Method (Phone)"; Status = "Failed"; Error = $_.Exception.Message }
                            $Result.Errors  += $_.Exception.Message
                            $Result.MFAResults.Failed += "Phone"; $failedCount++
                        }
                    }
 
                    "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                        try {
                            $null = Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $UserId -MicrosoftAuthenticatorAuthenticationMethodId $methodId -ErrorAction Stop
                            $Result.Actions += @{ Step = "Revoke MFA Method (Microsoft Authenticator)"; Status = "Success"; Error = $null }
                            $Result.MFAResults.Success += "Microsoft Authenticator"; $revokedCount++
                        }
                        catch {
                            $Result.Actions += @{ Step = "Revoke MFA Method (Microsoft Authenticator)"; Status = "Failed"; Error = $_.Exception.Message }
                            $Result.Errors  += $_.Exception.Message
                            $Result.MFAResults.Failed += "Microsoft Authenticator"; $failedCount++
                        }
                    }
 
                    "#microsoft.graph.softwareOathAuthenticationMethod" {
                        try {
                            $null = Remove-MgUserAuthenticationSoftwareOathMethod -UserId $UserId -SoftwareOathAuthenticationMethodId $methodId -ErrorAction Stop
                            $Result.Actions += @{ Step = "Revoke MFA Method (Software OATH)"; Status = "Success"; Error = $null }
                            $Result.MFAResults.Success += "Software OATH"; $revokedCount++
                        }
                        catch {
                            $Result.Actions += @{ Step = "Revoke MFA Method (Software OATH)"; Status = "Failed"; Error = $_.Exception.Message }
                            $Result.Errors  += $_.Exception.Message
                            $Result.MFAResults.Failed += "Software OATH"; $failedCount++
                        }
                    }
 
                    "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                        try {
                            $null = Remove-MgUserAuthenticationWindowsHelloForBusinessMethod -UserId $UserId -WindowsHelloForBusinessAuthenticationMethodId $methodId -ErrorAction Stop
                            $Result.Actions += @{ Step = "Revoke MFA Method (Windows Hello for Business)"; Status = "Success"; Error = $null }
                            $Result.MFAResults.Success += "Windows Hello for Business"; $revokedCount++
                        }
                        catch {
                            $Result.Actions += @{ Step = "Revoke MFA Method (Windows Hello for Business)"; Status = "Failed"; Error = $_.Exception.Message }
                            $Result.Errors  += $_.Exception.Message
                            $Result.MFAResults.Failed += "Windows Hello for Business"; $failedCount++
                        }
                    }
 
                    "#microsoft.graph.temporaryAccessPassAuthenticationMethod" {
                        try {
                            $null = Remove-MgUserAuthenticationTemporaryAccessPassMethod -UserId $UserId -TemporaryAccessPassAuthenticationMethodId $methodId -ErrorAction Stop
                            $Result.Actions += @{ Step = "Revoke MFA Method (Temporary Access Pass)"; Status = "Success"; Error = $null }
                            $Result.MFAResults.Success += "Temporary Access Pass"; $revokedCount++
                        }
                        catch {
                            $Result.Actions += @{ Step = "Revoke MFA Method (Temporary Access Pass)"; Status = "Failed"; Error = $_.Exception.Message }
                            $Result.Errors  += $_.Exception.Message
                            $Result.MFAResults.Failed += "Temporary Access Pass"; $failedCount++
                        }
                    }
 
                    Default {
                        $Result.Actions += @{ Step = "Revoke MFA Method ($methodType)"; Status = "Skipped"; Error = "Not supported" }
                        $Result.MFAResults.Skipped += $methodType
                        $skippedCount++
                    }
                }
            }
 
            $Result.Actions += @{
                Step   = "MFA Revocation Summary"
                Status = "Completed"
                Error  = $null
                Detail = @{
                    TotalMethods = $Methods.Count
                    Revoked      = $revokedCount
                    Skipped      = $skippedCount
                    Failed       = $failedCount
                }
            }
        }
    }
    catch {
        $Result.Actions += @{ Step = "Revoke MFA Methods"; Status = "Failed"; Error = $_.Exception.Message }
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
 
    # ─────────────────────────────────────────────
    # BUILD FILTERED JSON OUTPUT
    # ─────────────────────────────────────────────
 
    $FilteredSteps = @(
        "Authentication",
        "Disable Cloud Account",
        "Disable Primary On-prem Account",
        "Reset On-prem Password",
        "Disable Additional Domain Account",
        "Revoke Sign-In Sessions",
        "MFA Revocation Summary"
    )
 
    $FilteredActions = $Result.Actions | Where-Object {
        $step = $_.Step
        $FilteredSteps | ForEach-Object { if ($step -like "$_*") { return $true } }
    }
 
    $RevokeObj = $FilteredActions | Where-Object { $_.Step -eq "Revoke Sign-In Sessions" }
    $RevokeSessionsSummary = if ($RevokeObj) {
        @{ Status = $RevokeObj.Status; Error = $RevokeObj.Error }
    } else {
        @{ Status = "Skipped"; Error = "Step not executed" }
    }
 
    $MfaSummaryObj = $FilteredActions | Where-Object { $_.Step -eq "MFA Revocation Summary" }
    $MfaRevocationSummary = if ($MfaSummaryObj) {
        @{
            TotalMethods = $MfaSummaryObj.Detail.TotalMethods
            Revoked      = $MfaSummaryObj.Detail.Revoked
            Skipped      = $MfaSummaryObj.Detail.Skipped
            Failed       = $MfaSummaryObj.Detail.Failed
            SuccessList  = $Result.MFAResults.Success
            SkippedList  = $Result.MFAResults.Skipped
            FailedList   = $Result.MFAResults.Failed
        }
    } else {
        @{ TotalMethods = 0; Revoked = 0; Skipped = 0; Failed = 0; SuccessList = @(); SkippedList = @(); FailedList = @() }
    }
 
    $FilteredResult = @{
        ScriptName = $Result.ScriptName
        User       = $Result.User
        TimeStamp  = $Result.TimeStamp
        Status     = $Result.Status
        Summary    = @{
            Authentication = @{
                Status = ($FilteredActions | Where-Object { $_.Step -eq "Authentication" }).Status
                Error  = ($FilteredActions | Where-Object { $_.Step -eq "Authentication" }).Error
            }
            DisableCloudAccount = @{
                Status = ($FilteredActions | Where-Object { $_.Step -like "Disable*Cloud*" }).Status
                Error  = ($FilteredActions | Where-Object { $_.Step -like "Disable*Cloud*" }).Error
            }
            DisablePrimaryOnPrem = @{
                Status = ($FilteredActions | Where-Object { $_.Step -eq "Disable Primary On-prem Account" }).Status
                Error  = ($FilteredActions | Where-Object { $_.Step -eq "Disable Primary On-prem Account" }).Error
            }
            ResetOnPremPassword = @{
                Status = ($FilteredActions | Where-Object { $_.Step -eq "Reset On-prem Password" }).Status
                Error  = ($FilteredActions | Where-Object { $_.Step -eq "Reset On-prem Password" }).Error
            }
            DisableAdditionalOnPrem = @(
                $FilteredActions | Where-Object { $_.Step -like "Disable Additional Domain Account*" } | ForEach-Object {
                    @{
                        Domain = $_.Step -replace "Disable Additional Domain Account \(|\)", ""
                        Status = $_.Status
                        Error  = $_.Error
                    }
                }
            )
            RevokeSessions       = $RevokeSessionsSummary
            MFARevocationSummary = $MfaRevocationSummary
        }
    }
 
    $FilteredResult | ConvertTo-Json -Depth 5
}