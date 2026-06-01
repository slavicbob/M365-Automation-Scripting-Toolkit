<#
.SYNOPSIS
    LegalHoldAutomation.ps1
 
.DESCRIPTION
    This PowerShell script automates the process of:
    - Installing and importing the ExchangeOnlineManagement 3.6.0 module (if not already present).
    - Connecting to Exchange Online and Microsoft Purview Compliance Center (IPPSSession) using
      an admin service account with credential-based authentication.
    - Connecting to SharePoint Online (SPO) for OneDrive URL validation.
    - Verifying the existence of the specified eDiscovery case.
    - Checking if the specified mailbox exists and whether it already has an active InPlaceHold.
    - If no hold is present, creating a new Case Hold Policy and assigning it to the mailbox.
    - Optionally extending the hold to the mailbox owner's OneDrive site.
    - Enabling the hold and validating that it was successfully applied by checking InPlaceHolds
      after a delay.
    - Recording all actions, results, and exceptions in a structured $result object output as JSON.
 
.PARAMETER $caseId
    The unique GUID identifier of the eDiscovery case to associate with the new hold policy.
 
.PARAMETER $mailboxEmail
    The primary SMTP address of the mailbox to apply the legal hold on.
 
.PARAMETER Output
    A detailed $result object containing:
    - Module installation and import results
    - Exchange Online, IPPSSession, and SPO connection status
    - eDiscovery case verification status
    - Mailbox discovery and existing hold detection
    - OneDrive site URL lookup status
    - Hold policy creation, activation, and post-validation status
    - All exceptions and success messages from each operation
 
.EXAMPLE
    # Set credentials and target details, then run:
    .\LegalHoldAutomation.ps1
 
.NOTES
    - Requires the ExchangeOnlineManagement module version 3.6.0.
    - Requires the Microsoft.Online.SharePoint.PowerShell module.
    - Requires that the admin service account have the following roles:
        • Compliance Administrator
        • Exchange Administrator
        • Security Administrator
        • SharePoint Administrator
    - InPlaceHold validation uses an iterative loop-based delay (15s x 6 retries) to allow
      backend propagation.
    - Script is suitable for automation pipelines and audit trails due to JSON-based result logging.
 
.ChangeLog
    V1.00 - Initial version
#>
 
# ─────────────────────────────────────────────
# CONFIGURATION — replace all placeholder values
# ─────────────────────────────────────────────
 
# Admin service account credentials
$username   = "admin-service-account@<YOUR_DOMAIN>"
$password   = ConvertTo-SecureString "<YOUR_PASSWORD>" -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($username, $password)
$cred       = $credential
 
# SharePoint Online admin center URL
$spoAdminUrl = "https://<YOUR_TENANT>-admin.sharepoint.com/"
 
# eDiscovery case and target mailbox
$caseId          = "<YOUR_EDISCOVERY_CASE_GUID>"
$mailboxEmail    = "target.user@<YOUR_DOMAIN>"
$holdPolicyName  = "Hold-$($mailboxEmail.Split('@')[0])-$(Get-Date -Format 'yyyy-MM-dd')"
$description     = ""
 
# ─────────────────────────────────────────────
# MODULE SETUP
# ─────────────────────────────────────────────
 
$moduleName      = "ExchangeOnlineManagement"
$requiredVersion = "3.6.0"
 
$result = @{
    ModuleStatus = @{
        Install = @()
        Import  = @()
    }
    ConnectionStatus = @{
        ExchangeOnline          = @()
        IPPSSession             = @()
        DisconnectExchangeOnline = @()
        DisconnectIPPSSession   = @()
        SPOSession              = @()
        DisconnectSPOSession    = @()
    }
    eDiscoveryCase = @{
        CaseId   = @()
        CaseName = @()
        Status   = @()
    }
    Mailbox = @{
        Email        = @()
        Status       = @()
        InPlaceHolds = @()
    }
    SharePointSite = @{
        Status = @()
        URL    = @()
    }
    HoldPolicy = @{
        PolicyName = @()
        Status     = @()
        Error      = @()
    }
}
 
# ─────────────────────────────────────────────
# INSTALL / IMPORT ExchangeOnlineManagement
# ─────────────────────────────────────────────
 
try {
    if (-not (Get-InstalledModule -Name $moduleName -RequiredVersion $requiredVersion -ErrorAction SilentlyContinue)) {
        try {
            Install-Module -Name $moduleName -RequiredVersion $requiredVersion -Force -Scope CurrentUser -ErrorAction Stop
            $result.ModuleStatus.Install = "Installed $moduleName $requiredVersion successfully."
            Start-Sleep -Seconds 5
            Import-Module -Name $moduleName -RequiredVersion $requiredVersion -Force -ErrorAction Stop
            $result.ModuleStatus.Import = "Imported $moduleName $requiredVersion after installation."
        }
        catch {
            $result.ModuleStatus.Install = "Installation failed: $($_.Exception.Message)"
            $result.ModuleStatus.Import  = "Import skipped due to failed installation."
            $result | ConvertTo-Json -Depth 5
            return
        }
    }
    else {
        $result.ModuleStatus.Install = "$moduleName $requiredVersion already installed."
        try {
            Start-Sleep -Seconds 5
            Import-Module -Name $moduleName -RequiredVersion $requiredVersion -Force -ErrorAction Stop
            $result.ModuleStatus.Import = "Imported $moduleName $requiredVersion (already installed)."
        }
        catch {
            $result.ModuleStatus.Import = "Import failed: $($_.Exception.Message)"
            $result | ConvertTo-Json -Depth 5
            return
        }
    }
}
catch {
    $result.ModuleStatus.Install = "Fatal module setup error: $($_.Exception.Message)"
    $result | ConvertTo-Json -Depth 5
    return
}
 
# ─────────────────────────────────────────────
# CONNECT TO SHAREPOINT ONLINE
# ─────────────────────────────────────────────
 
try {
    $sposModuleName = "Microsoft.Online.SharePoint.PowerShell"
    if (-not (Get-Module -ListAvailable -Name $sposModuleName)) {
        Install-Module -Name $sposModuleName -Force -Scope CurrentUser -ErrorAction Stop
        $result.ConnectionStatus.SPOSession = "Installed $sposModuleName successfully."
    }
    else {
        $result.ConnectionStatus.SPOSession = "$sposModuleName already installed."
    }
    Import-Module -Name $sposModuleName -Force -ErrorAction Stop -WarningAction SilentlyContinue
    Connect-SPOService -Credential $credential -Url $spoAdminUrl -ErrorAction Stop -WarningAction SilentlyContinue
    $result.ConnectionStatus.SPOSession = "Connected successfully."
}
catch {
    $result.ConnectionStatus.SPOSession = "Connection failed: $($_.Exception.Message)"
    $result | ConvertTo-Json -Depth 5
    return
}
 
# ─────────────────────────────────────────────
# MAIN LOGIC
# ─────────────────────────────────────────────
 
try {
 
    # Connect to Exchange Online
    try {
        Connect-ExchangeOnline -Credential $cred -ShowBanner:$false
        $result.ConnectionStatus.ExchangeOnline = "Connected successfully."
    }
    catch {
        $result.ConnectionStatus.ExchangeOnline = "Connection failed: $($_.Exception.Message)"
        $result | ConvertTo-Json -Depth 5
        return
    }
 
    # Connect to Microsoft Purview (IPPSSession)
    try {
        Connect-IPPSSession -Credential $cred -ShowBanner:$false
        $result.ConnectionStatus.IPPSSession = "Connected successfully."
    }
    catch {
        $result.ConnectionStatus.IPPSSession = "Connection failed: $($_.Exception.Message)"
        $result | ConvertTo-Json -Depth 5
        return
    }
 
    # Verify eDiscovery case
    try {
        $case = Get-ComplianceCase -Identity $caseId -ErrorAction Stop
        $result.eDiscoveryCase.CaseId   = $caseId
        $result.eDiscoveryCase.CaseName = $case.Name
        $result.eDiscoveryCase.Status   = "Case found."
    }
    catch {
        $result.eDiscoveryCase.Status = "Case lookup failed: $($_.Exception.Message)"
        $result | ConvertTo-Json -Depth 5
        return
    }
 
    # Verify mailbox and check for existing holds
    try {
        $mailbox = Get-Mailbox -Identity $mailboxEmail -ErrorAction Stop | Select-Object PrimarySmtpAddress, InPlaceHolds
        $result.Mailbox.Email  = $mailbox.PrimarySmtpAddress.ToString()
        $result.Mailbox.Status = "Mailbox found."
        if ($mailbox.InPlaceHolds -and $mailbox.InPlaceHolds.Count -gt 0) {
            $result.Mailbox.InPlaceHolds = $mailbox.InPlaceHolds
            $result.Mailbox.Status       = "Hold already exists. Skipping hold creation."
            $result | ConvertTo-Json -Depth 5
            return
        }
        else {
            $result.Mailbox.InPlaceHolds = "No active holds."
        }
    }
    catch {
        $result.Mailbox.Status = "Mailbox lookup or InPlaceHolds check failed: $($_.Exception.Message)"
        $result | ConvertTo-Json -Depth 5
        return
    }
 
    # Look up OneDrive URL for the mailbox owner
    $oneDriveUrl = $null
    try {
        $filter      = "Owner -eq '$mailboxEmail'"
        $oneDriveUrl = (Get-SPOSite -Filter $filter -IncludePersonalSite $true -Limit All).Url
        if ($oneDriveUrl) {
            $result.SharePointSite.URL    = $oneDriveUrl
            $result.SharePointSite.Status = "OneDrive site found."
        }
        else {
            $result.SharePointSite.Status = "No OneDrive site found."
        }
    }
    catch {
        $result.SharePointSite.Status = "OneDrive site lookup failed: $($_.Exception.Message)"
        $result | ConvertTo-Json -Depth 5
        return
    }
 
    # Create hold policy and validate
    try {
        New-CaseHoldPolicy -Name $holdPolicyName -Case $caseId -ExchangeLocation $mailboxEmail -Enabled $true -Comment $description -ErrorAction Stop | Out-Null
        Start-Sleep -Seconds 10
        $result.HoldPolicy.PolicyName = $holdPolicyName
        $result.HoldPolicy.Status     = "Hold policy created successfully."
 
        if ($oneDriveUrl) {
            Set-CaseHoldPolicy -Identity $holdPolicyName -Enabled $true -AddSharePointLocation $oneDriveUrl -ErrorAction Stop
        }
        else {
            Set-CaseHoldPolicy -Identity $holdPolicyName -Enabled $true -ErrorAction Stop
        }
 
        # Poll for InPlaceHold propagation (up to 6 retries x 15s)
        $confirmed = $false
        $retries   = 0
        do {
            Start-Sleep -Seconds 15
            $confirmedMailbox = Get-Mailbox -Identity $mailboxEmail -ErrorAction Stop | Select-Object PrimarySmtpAddress, InPlaceHolds
            if ($confirmedMailbox.InPlaceHolds -and $confirmedMailbox.InPlaceHolds.Count -gt 0) {
                $result.Mailbox.InPlaceHolds  = $confirmedMailbox.InPlaceHolds
                $result.HoldPolicy.Status     = "Hold confirmed via InPlaceHolds after $retries retry(ies)."
                $confirmed = $true
            }
            $retries++
        } while (-not $confirmed -and $retries -lt 6)
 
        if (-not $confirmed) {
            $result.HoldPolicy.Status = "Hold creation did not reflect in InPlaceHolds after $retries retries."
        }
    }
    catch {
        $result.HoldPolicy.PolicyName = $holdPolicyName
        $result.HoldPolicy.Status     = "Creation or confirmation failed."
        $result.HoldPolicy.Error      = $_.Exception.Message
        $result | ConvertTo-Json -Depth 5
        return
    }
 
}
finally {
 
    # Disconnect Exchange Online + IPPSSession
    try {
        Disconnect-ExchangeOnline -Confirm:$false
        $result.ConnectionStatus.DisconnectExchangeOnline = "Disconnected successfully."
        $result.ConnectionStatus.DisconnectIPPSSession    = "Disconnected successfully."
    }
    catch {
        $result.ConnectionStatus.DisconnectExchangeOnline = "Disconnection failed: $($_.Exception.Message)"
        $result.ConnectionStatus.DisconnectIPPSSession    = "Disconnection failed: $($_.Exception.Message)"
    }
 
    # Disconnect SharePoint Online
    try {
        Disconnect-SPOService
        $result.ConnectionStatus.DisconnectSPOSession = "Disconnected from SharePoint Online."
    }
    catch {
        $result.ConnectionStatus.DisconnectSPOSession = "SharePoint disconnection failed: $($_.Exception.Message)"
    }
 
}
 
$result | ConvertTo-Json -Depth 5