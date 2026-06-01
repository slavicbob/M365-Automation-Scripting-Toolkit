<#
.SYNOPSIS
    Creates and configures an Exchange Online distribution list from a ServiceNow catalog request.
 
.DESCRIPTION
    This PowerShell script automates the provisioning of a new Exchange Online distribution list (DL)
    based on input parameters passed from a ServiceNow catalog item. It performs the following:
 
    - Connecting to Exchange Online using certificate-based app authentication.
    - Creating a new Distribution Group with a generated primary SMTP address.
    - Adding all specified members to the group.
    - Configuring whether the DL accepts external (unauthenticated) email.
    - Configuring Global Address Book visibility (Show or Hide).
    - Outputting a structured JSON result object logging every action and any errors.
 
.PARAMETER DLName
    The name of the distribution list to create. Also used to derive the primary SMTP address.
 
.PARAMETER Members
    An array of member email addresses to add to the distribution list upon creation.
 
.PARAMETER ExternalEmailSetting
    Controls whether external senders can email the DL.
    Accepted values: "Open" (allow external) or "Closed" (block external).
 
.PARAMETER AddressBookVisibility
    Controls whether the DL appears in the Global Address Book.
    Accepted values: "Show" or "Hide".
 
.OUTPUTS
    A JSON result object containing connection, creation, member addition, setting, and
    disconnection status for each step, plus an aggregated errors array.
 
.EXAMPLE
    .\DistributionGroup_Creation.ps1 `
        -DLName "team-finance" `
        -Members @("user1@<YOUR_DOMAIN>", "user2@<YOUR_DOMAIN>") `
        -ExternalEmailSetting "Open" `
        -AddressBookVisibility "Show"
 
.NOTES
    Required permissions on the Entra ID app registration:
        - Exchange.ManageAsApp (Exchange Online)
 
    The service principal must be assigned the Exchange Administrator role or a
    scoped Exchange management role that permits distribution group creation.
 
.ChangeLog
    V1.00 - Initial version
#>
 
# Input parameters from ServiceNow catalog item
param (
    [Parameter(Mandatory = $true)]
    [string]$DLName,                    # Distribution List Name
 
    [Parameter(Mandatory = $true)]
    [string[]]$Members,                 # Array of member email addresses
 
    [Parameter(Mandatory = $true)]
    [ValidateSet("Open", "Closed")]
    [string]$ExternalEmailSetting,      # External email permission (Open or Closed)
 
    [Parameter(Mandatory = $true)]
    [ValidateSet("Show", "Hide")]
    [string]$AddressBookVisibility      # Show or Hide in Global Address Book
)
 
# ─────────────────────────────────────────────
# CONFIGURATION — replace all placeholder values
# ─────────────────────────────────────────────
 
$AppId      = "<YOUR_CLIENT_ID>"
$Thumbprint = "<YOUR_CERT_THUMBPRINT>"
$domainName = "<YOUR_DOMAIN>"
 
# Exchange Online organization (e.g. contoso.onmicrosoft.com or your verified domain)
$Organization = "<YOUR_EXCHANGE_ORGANIZATION>"
 
$primarysmtpaddress = "$DLName@$domainName"
 
# ─────────────────────────────────────────────
# INITIALIZE RESULT LOG
# ─────────────────────────────────────────────
 
$result = @{
    Timestamp             = (Get-Date).ToString("s")
    DLName                = $DLName
    PrimarySmtpAddress    = $primarysmtpaddress
    Status                = "Started"
    Connection            = @{}
    Creation              = @{}
    MembersAddition       = @()
    ExternalEmailSetting  = @{}
    AddressBookVisibility = @{}
    Disconnection         = @{}
    Errors                = @()
}
 
try {
 
    # ─────────────────────────────────────────────
    # CONNECT TO EXCHANGE ONLINE
    # ─────────────────────────────────────────────
 
    try {
        Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $Thumbprint -Organization $Organization -ErrorAction Stop -ShowBanner:$false
        $result.Connection = @{
            Time    = (Get-Date).ToString("s")
            Status  = "Success"
            Message = "Connected to Exchange Online."
        }
    }
    catch {
        $result.Connection = @{
            Time    = (Get-Date).ToString("s")
            Status  = "Failed"
            Message = $_.Exception.Message
        }
        throw $_
    }
 
    # ─────────────────────────────────────────────
    # CREATE DISTRIBUTION LIST
    # ─────────────────────────────────────────────
 
    try {
        New-DistributionGroup -Name $DLName -Type Distribution -PrimarySmtpAddress $primarysmtpaddress -ErrorAction Stop
        $result.Creation = @{
            Time    = (Get-Date).ToString("s")
            Status  = "Success"
            Message = "Distribution list created."
        }
    }
    catch {
        $result.Creation = @{
            Time    = (Get-Date).ToString("s")
            Status  = "Failed"
            Message = $_.Exception.Message
        }
        throw $_
    }
 
    # ─────────────────────────────────────────────
    # ADD MEMBERS
    # ─────────────────────────────────────────────
 
    foreach ($Member in $Members) {
        $memberResult = @{
            Member  = $Member
            Time    = (Get-Date).ToString("s")
            Status  = ""
            Message = ""
        }
        try {
            Add-DistributionGroupMember -Identity $DLName -Member $Member -ErrorAction Stop
            $memberResult.Status  = "Success"
            $memberResult.Message = "Member added."
        }
        catch {
            $memberResult.Status  = "Failed"
            $memberResult.Message = $_.Exception.Message
            $result.Errors       += $_.Exception.Message
        }
        $result.MembersAddition += $memberResult
    }
 
    # ─────────────────────────────────────────────
    # EXTERNAL EMAIL SETTING
    # ─────────────────────────────────────────────
 
    try {
        $authSetting = if ($ExternalEmailSetting -eq "Open") { $false } else { $true }
        Set-DistributionGroup -Identity $DLName -RequireSenderAuthenticationEnabled $authSetting -ErrorAction Stop
        $result.ExternalEmailSetting = @{
            Time    = (Get-Date).ToString("s")
            Status  = "Success"
            Setting = $ExternalEmailSetting
            Message = "External email setting applied."
        }
    }
    catch {
        $result.ExternalEmailSetting = @{
            Time    = (Get-Date).ToString("s")
            Status  = "Failed"
            Setting = $ExternalEmailSetting
            Message = $_.Exception.Message
        }
        $result.Errors += $_.Exception.Message
        throw $_
    }
 
    # ─────────────────────────────────────────────
    # ADDRESS BOOK VISIBILITY
    # ─────────────────────────────────────────────
 
    try {
        $visibility = if ($AddressBookVisibility -eq "Hide") { $true } else { $false }
        Set-DistributionGroup -Identity $DLName -HiddenFromAddressListsEnabled $visibility -ErrorAction Stop
        $result.AddressBookVisibility = @{
            Time    = (Get-Date).ToString("s")
            Status  = "Success"
            Setting = $AddressBookVisibility
            Message = "Address book visibility configured."
        }
    }
    catch {
        $result.AddressBookVisibility = @{
            Time    = (Get-Date).ToString("s")
            Status  = "Failed"
            Setting = $AddressBookVisibility
            Message = $_.Exception.Message
        }
        $result.Errors += $_.Exception.Message
        throw $_
    }
 
    $result.Status = "Completed Successfully"
}
catch {
    $result.Status  = "Failed"
    $result.Errors += $_.Exception.Message
}
finally {
    try {
        Disconnect-ExchangeOnline -Confirm:$false
        $result.Disconnection = @{
            Time    = (Get-Date).ToString("s")
            Status  = "Success"
            Message = "Disconnected from Exchange Online."
        }
    }
    catch {
        $result.Disconnection = @{
            Time    = (Get-Date).ToString("s")
            Status  = "Failed"
            Message = $_.Exception.Message
        }
        $result.Errors += $_.Exception.Message
    }
 
    $result | ConvertTo-Json -Depth 5 | Write-Output
}
 