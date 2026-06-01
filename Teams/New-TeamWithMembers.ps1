<#
.SYNOPSIS
    New-TeamWithMembers.ps1 — Creates a Microsoft Teams team with owners and members via Microsoft Graph.
 
.DESCRIPTION
    This PowerShell script automates the full provisioning of a Microsoft Teams team using the
    Microsoft Graph API. It performs the following steps in sequence:
 
    - Installing and importing the Microsoft.Graph module if not already present.
    - Authenticating to Microsoft Graph via certificate-based app authentication.
    - Creating a Microsoft 365 Unified Group with a generated mail nickname.
    - Adding specified users as owners of the group.
    - Adding specified users as members of the group.
    - Provisioning a Teams layer on top of the group (PUT /groups/{id}/team) with
      configurable messaging, member, and fun settings.
    - Outputting a structured JSON result object logging every action taken.
 
.OUTPUTS
    A JSON result object containing step-by-step details, overall status, start/end timestamps,
    and any error messages encountered during execution.
 
.EXAMPLE
    # Set configuration values in the CONFIGURATION block, then run:
    .\New-TeamWithMembers.ps1
 
.NOTES
    Required Entra ID app registration permissions (application):
        - Group.ReadWrite.All
        - TeamSettings.ReadWrite.All
        - User.Read.All
 
    The Graph beta endpoint is used for the final team provisioning step (PUT /groups/{id}/team).
 
.ChangeLog
    V1.00 - Initial version
#>
 
# ─────────────────────────────────────────────
# CONFIGURATION — replace all placeholder values
# ─────────────────────────────────────────────
 
$ClientId    = "<YOUR_CLIENT_ID>"
$TenantId    = "<YOUR_TENANT_ID>"
$Thumbprint  = "<YOUR_CERT_THUMBPRINT>"
 
$teamDisplayName = "<YOUR_TEAM_DISPLAY_NAME>"
$teamDescription = "<YOUR_TEAM_DESCRIPTION>"
$teamVisibility  = "Public"   # Options: "Private", "Public"
 
$owners  = @("owner1@<YOUR_DOMAIN>", "owner2@<YOUR_DOMAIN>")
$members = @("member1@<YOUR_DOMAIN>", "member2@<YOUR_DOMAIN>")
 
# ─────────────────────────────────────────────
# INITIALIZE LOGGING OBJECT
# ─────────────────────────────────────────────
 
$result = [ordered]@{
    Script           = "New-TeamWithMembers.ps1"
    StartTime        = (Get-Date)
    TeamDisplayName  = $teamDisplayName
    RequestedOwners  = $owners
    RequestedMembers = $members
    Status           = "Started"
    Details          = @()
    ModuleStatus     = @()
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
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -ErrorAction Stop -NoWelcome
        $result.Details += "Connected to Microsoft Graph."
    }
    catch {
        $result.Status   = "Failed"
        $result.Details += "Graph Connection Failed: $($_.Exception.Message)"
        throw
    }
 
    # ─────────────────────────────────────────────
    # CREATE MICROSOFT 365 GROUP
    # ─────────────────────────────────────────────
 
    try {
        $mailNick = ($teamDisplayName -replace '[^a-zA-Z0-9]', '') + (Get-Random -Maximum 9999)
        $group = New-MgGroup `
            -DisplayName $teamDisplayName `
            -Description $teamDescription `
            -MailEnabled:$true `
            -MailNickname $mailNick `
            -SecurityEnabled:$false `
            -GroupTypes "Unified" `
            -Visibility $teamVisibility `
            -ErrorAction Stop
        $result.Details += "Group created with ID: $($group.Id)"
    }
    catch {
        $result.Status   = "Failed"
        $result.Details += "Group Creation Failed: $($_.Exception.Message)"
        throw
    }
 
    # ─────────────────────────────────────────────
    # ADD OWNERS
    # ─────────────────────────────────────────────
 
    foreach ($owner in $owners) {
        try {
            $user = Get-MgUser -UserId $owner -ErrorAction Stop
            New-MgGroupOwnerByRef -GroupId $group.Id -BodyParameter @{
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
    # ADD MEMBERS
    # ─────────────────────────────────────────────
 
    foreach ($member in $members) {
        try {
            $user = Get-MgUser -UserId $member -ErrorAction Stop
            New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter @{
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
    # PROVISION TEAMS LAYER ON GROUP
    # ─────────────────────────────────────────────
 
    try {
        $teamPayload = @{
            memberSettings = @{
                allowCreateUpdateChannels = $true
                allowDeleteChannels       = $true
            }
            messagingSettings = @{
                allowUserEditMessages   = $true
                allowUserDeleteMessages = $true
            }
            funSettings = @{
                allowGiphy         = $true
                giphyContentRating = "strict"
            }
        } | ConvertTo-Json -Depth 5
 
        Invoke-MgGraphRequest -Method PUT `
            -Uri "https://graph.microsoft.com/beta/groups/$($group.Id)/team" `
            -Body $teamPayload -ContentType "application/json" -ErrorAction Stop
 
        $result.Details += "Team created successfully on group."
    }
    catch {
        $result.Status   = "Failed"
        $result.Details += "Team Creation Failed: $($_.Exception.Message)"
        throw
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
