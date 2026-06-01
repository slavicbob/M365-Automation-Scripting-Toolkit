<#
.SYNOPSIS
    Capture-AndUpload-AutopilotHash.ps1
    Run as SYSTEM via Intune Platform Script / Remediation. No parameters.
 
.SYNOPSIS
    Captures Autopilot hardware hash and uploads as a CSV that exactly matches
    the format produced by Get-WindowsAutopilotInfo.ps1, to a SharePoint
    Documents library via Microsoft Graph (Sites.Selected).
 
.NOTES
    - Must run as SYSTEM (required for the MDM_DevDetail_Ext01 WMI namespace).
    - Designed for Intune Platform Script or Remediation deployment.
    - Output format matches Microsoft's official script byte-for-byte.
#>
 
$ErrorActionPreference = 'Stop'
 
# ============== CONFIG ==============
# Replace all values below before deployment.
 
$tenantId     = "<YOUR_TENANT_ID>"
$clientId     = "<YOUR_CLIENT_ID>"
$clientSecret = "<YOUR_CLIENT_SECRET>"
 
# SharePoint site ID — format: <tenant>.sharepoint.com,<site-guid>,<web-guid>
$siteId       = "<YOUR_TENANT>.sharepoint.com,<YOUR_SITE_GUID>,<YOUR_WEB_GUID>"
 
# Target folder path within the site's default document library
$folderPath   = "HardwareHash"
 
# Optional — matches -GroupTag in Get-WindowsAutopilotInfo.ps1
$groupTag     = ""
 
# Optional — matches -AssignedUser in Get-WindowsAutopilotInfo.ps1
$assignedUser = ""
 
# ====================================
 
# Logging
$logDir  = Join-Path $env:ProgramData 'AutopilotHashCapture'
$logFile = Join-Path $logDir 'capture.log'
if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
 
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    "$([datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')) [$Level] $Message" |
        Out-File -FilePath $logFile -Append -Encoding utf8
}
 
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
 
try {
    Write-Log "=== Starting hash capture on $env:COMPUTERNAME ==="
 
    # 1. Acquire Graph token
    $tokenResp = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body @{
            grant_type    = 'client_credentials'
            client_id     = $clientId
            client_secret = $clientSecret
            scope         = 'https://graph.microsoft.com/.default'
        }
    $headers = @{ Authorization = "Bearer $($tokenResp.access_token)" }
 
    # 2. Capture hardware hash + serial (same WMI logic as Get-WindowsAutopilotInfo)
    $session   = New-CimSession
    $devDetail = Get-CimInstance -CimSession $session `
                    -Namespace 'root/cimv2/mdm/dmmap' `
                    -ClassName 'MDM_DevDetail_Ext01' `
                    -Filter "InstanceID='Ext' AND ParentID='./DevDetail'"
    $bios = Get-CimInstance -CimSession $session -ClassName Win32_BIOS
    Remove-CimSession $session
 
    $serial = ($bios.SerialNumber | Out-String).Trim()
    $hash   = $devDetail.DeviceHardwareData
 
    if ([string]::IsNullOrWhiteSpace($serial)) { throw 'Serial number is empty.' }
    if ([string]::IsNullOrWhiteSpace($hash))   { throw 'Hardware hash is empty - script must run as SYSTEM.' }
 
    Write-Log "Serial: $serial | Hash length: $($hash.Length) chars"
    $safeSerial = ($serial -replace '[\\/:\*\?"<>\|]', '_').Trim()
 
    # 3. Ensure destination folder exists
    $folderUri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/root:/$folderPath"
    try {
        Invoke-RestMethod -Method Get -Uri $folderUri -Headers $headers | Out-Null
    }
    catch {
        if ($_.Exception.Response.StatusCode.value__ -eq 404) {
            $folderBody = @{
                name = $folderPath; folder = @{}
                '@microsoft.graph.conflictBehavior' = 'replace'
            } | ConvertTo-Json
            Invoke-RestMethod -Method Post `
                -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drive/root/children" `
                -Headers $headers -ContentType 'application/json' -Body $folderBody | Out-Null
        }
        else { throw }
    }
 
    # 4. Build the CSV exactly the way Get-WindowsAutopilotInfo.ps1 does.
    #    Property order matters - this is the order the official script uses.
    $computer = New-Object -TypeName PSObject -Property ([ordered]@{
        'Device Serial Number' = $serial
        'Windows Product ID'   = ''        # Always blank in the official script
        'Hardware Hash'        = $hash
    })
 
    # ConvertTo-Csv -NoTypeInformation produces the same output as Export-Csv,
    # which is what Get-WindowsAutopilotInfo uses internally.
    $csvLines = $computer | ConvertTo-Csv -NoTypeInformation
    $csvText  = ($csvLines -join "`r`n") + "`r`n"
 
    # 5. Upload as <serial>.csv
    $fileName  = "$safeSerial.csv"
    $uploadUri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/root:/$folderPath/$fileName" + ':/content'
    Write-Log "Uploading $fileName"
 
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($csvText)
    Invoke-RestMethod -Method Put -Uri $uploadUri `
        -Headers $headers -ContentType 'text/csv' -Body $bytes | Out-Null
 
    Write-Log "=== Upload successful for $serial ==="
    Write-Output "Uploaded hash for $serial"
}
catch {
    $errMsg = "Capture/upload failed: $($_.Exception.Message)"
    Write-Log $errMsg 'ERROR'
    Write-Error $errMsg
}