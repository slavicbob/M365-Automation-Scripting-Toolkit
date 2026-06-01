<#
.SYNOPSIS
    Update-TeamMembersOwners.ps1 — Adds and removes owners and members from an existing Microsoft Teams team.
 
.DESCRIPTION
    This PowerShell script manages membership of an existing Microsoft Teams team (via its
    underlying Microsoft 365 Group) using Microsoft Graph. It performs the following:
 
    - Installing and importing the Microsoft.Graph module if not already present.
    - Authenticating to Microsoft Graph via certificate-based app authentication.
    - Adding specified users as group owners.
    - Removing specified users from group owners.
    - Adding specified users as group members.
    - Removing specified users from group members.
    - Outputting a structured JSON result object logging every action taken.
 
.OUTPUTS
    A JSON result object with step-by-step details, overall status, and start/end timestamps.
 
.EXAMPLE
    # Set configuration values in the CONFIGURATION block, then run:
    .\Update-TeamMembersOwners.ps1
 
.NOTES
    Required Entra ID app registration permissions (application):
        - Group.ReadWrite.All
        - User.Read.All
 
    The $teamId parameter targets the underlying Microsoft 365 Group ID of the team,
    which is the same as the Teams team ID.
 
.ChangeLog
    V1.00 - Initial version
#>
 
# ─────────────────────────────────────────────
# CONFIGURATION — replace all placeholder values
# ─────────────────────────────────────────────
 
$clientId   = "<YOUR_CLIENT_ID>"
$tenantId   = "<YOUR_TENANT_ID>"
$thumbprint = "<YOUR_CERT_THUMBPRINT>"
 
# Target Teams team ID (same as the underlying M365 Group ID)
$teamId = "<YOUR_TEAM_ID>"
 
# Owners to add/remove
$ownersToAdd    = @("owner1@<YOUR_DOMAIN>")
$ownersToRemove = @("owner2@<YOUR_DOMAIN>")
 
# Members to add/remove
$membersToAdd    = @("member1@<YOUR_DOMAIN>")
$membersToRemove = @("member2@<YOUR_DOMAIN>")
 
# ─────────────────────────────────────────────
# INITIALIZE LOGGING OBJECT
# ─────────────────────────────────────────────
 
$result = [ordered]@{
    Script       = "Update-TeamMembersOwners.ps1"
    StartTime    = (Get-Date)
    TeamId       = $teamId
    Status       = "Started"
    Details      = @()
    ModuleStatus = @()
}
 
try {
 
    # ─────────────────────────────────────────────
    # INSTALL / IMPORT MICROSOFT.GRAPH
    # ─────────────────────────────────────────────
 
    foreach ($module in @("Microsoft.Graph")) {
        try {
            if (-not (Get-Module -ListAvailable -Name $module)) {
                Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                $result.ModuleStatus += "$module installed successfully."
            }
            else {
                $result.ModuleStatus += "$module already installed."
            }
            $result.ModuleStatus += "$module imported successfully."
        }
        catch {
            $result.ModuleStatus += "$module error: $($_.Exception.Message)"
            throw
        }
    }
 
    # ─────────────────────────────────────────────
    # CONNECT TO MICROSOFT GRAPH
    # ─────────────────────────────────────────────
 
    try {
        Connect-MgGraph -TenantId $tenantId -ClientId $clientId -CertificateThumbprint $thumbprint -ErrorAction Stop -NoWelcome
        $result.Details += "Connected to Microsoft Graph."
    }
    catch {
        $result.Status   = "Failed"
        $result.Details += "Graph Connection Failed: $($_.Exception.Message)"
        throw
    }
 
    # ─────────────────────────────────────────────
    # ADD OWNERS
    # ─────────────────────────────────────────────
 
    foreach ($owner in $ownersToAdd) {
        try {
            $user = Get-MgUser -UserId $owner -ErrorAction Stop
            New-MgGroupOwnerByRef -GroupId $teamId -BodyParameter @{
                "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($user.Id)"
            } -ErrorAction Stop
            $result.Details += "Added owner: $owner"
        }
        catch {
            $result.Status   = "Failed"
            $result.Details += "Failed to add owner '$owner': $($_.Exception.Message)"
            throw
        }
    }
 
    # ─────────────────────────────────────────────
    # REMOVE OWNERS
    # ─────────────────────────────────────────────
 
    foreach ($owner in $ownersToRemove) {
        try {
            $user = Get-MgUser -UserId $owner -ErrorAction Stop
            Remove-MgGroupOwnerByRef -GroupId $teamId -DirectoryObjectId $user.Id -ErrorAction Stop
            $result.Details += "Removed owner: $owner"
        }
        catch {
            $result.Status   = "Failed"
            $result.Details += "Failed to remove owner '$owner': $($_.Exception.Message)"
            throw
        }
    }
 
    # ─────────────────────────────────────────────
    # ADD MEMBERS
    # ─────────────────────────────────────────────
 
    foreach ($member in $membersToAdd) {
        try {
            $user = Get-MgUser -UserId $member -ErrorAction Stop
            New-MgGroupMemberByRef -GroupId $teamId -BodyParameter @{
                "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($user.Id)"
            } -ErrorAction Stop
            $result.Details += "Added member: $member"
        }
        catch {
            $result.Status   = "Failed"
            $result.Details += "Failed to add member '$member': $($_.Exception.Message)"
            throw
        }
    }
 
    # ─────────────────────────────────────────────
    # REMOVE MEMBERS
    # ─────────────────────────────────────────────
 
    foreach ($member in $membersToRemove) {
        try {
            $user = Get-MgUser -UserId $member -ErrorAction Stop
            Remove-MgGroupMemberByRef -GroupId $teamId -DirectoryObjectId $user.Id -ErrorAction Stop
            $result.Details += "Removed member: $member"
        }
        catch {
            $result.Status   = "Failed"
            $result.Details += "Failed to remove member '$member': $($_.Exception.Message)"
            throw
        }
    }
 
    $result.Status = "Success"
}
catch {
    # Errors are logged in individual step blocks above
}
finally {
    Disconnect-MgGraph | Out-Null
    $result.EndTime = (Get-Date)
    $result | ConvertTo-Json -Depth 5
}