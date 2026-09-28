<#
.SYNOPSIS
    Downloads the full OneDrive contents of an Active Directory user to C:\Temp\OneDriveArchive.

.DESCRIPTION
    1. Prompts for an AD username (sAMAccountName, e.g. jsmith).
    2. Looks the user up in Active Directory to get their UserPrincipalName.
    3. Connects to Microsoft Graph and locates the user's OneDrive.
    4. Recursively downloads every file, preserving the folder structure, to
       C:\Temp\OneDriveArchive\<username>\
    5. Writes a log file (and a CSV of any failures) next to the archive.

    Re-running the script for the same user resumes: files that already exist
    locally with the same size are skipped.

.REQUIREMENTS
    - RSAT ActiveDirectory module   (Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0)
    - Microsoft Graph PowerShell    (Install-Module Microsoft.Graph.Authentication -Scope CurrentUser)
    - Permissions, one of:
        * App-only (recommended): an Entra ID app registration with the
          *Application* permissions Files.Read.All and User.Read.All (admin
          consented) and a certificate. Pass -TenantId, -ClientId and
          -CertificateThumbprint.
        * Delegated (interactive sign-in): Files.Read.All + User.Read.All, AND
          the signed-in admin must have access to the target user's OneDrive
          (e.g. added as Site Collection Admin in the SharePoint admin center).

.EXAMPLE
    .\Export-UserOneDrive.ps1
    Prompts for the username and signs in interactively.

.EXAMPLE
    .\Export-UserOneDrive.ps1 -TenantId contoso.onmicrosoft.com -ClientId <appId> -CertificateThumbprint <thumbprint>
    Uses app-only authentication.
#>
[CmdletBinding()]
param(
    [string]$DestinationRoot = 'C:\Temp\OneDriveArchive',
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

$ErrorActionPreference = 'Stop'

#region Helpers ---------------------------------------------------------------

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    switch ($Level) {
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
    if ($script:LogFile) { Add-Content -Path $script:LogFile -Value $line }
}

# Strip characters Windows does not allow in file/folder names.
function Get-SafeName {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars() -join ''
    $safe = $Name -replace "[$([regex]::Escape($invalid))]", '_'
    return $safe.TrimEnd('.', ' ')
}

# Invoke-MgGraphRequest with simple retry for throttling (429) / transient 5xx.
function Invoke-GraphWithRetry {
    param([string]$Uri, [string]$OutputFilePath, [int]$MaxAttempts = 5)
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            if ($OutputFilePath) {
                return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputFilePath $OutputFilePath
            }
            return Invoke-MgGraphRequest -Method GET -Uri $Uri
        }
        catch {
            $status = $_.Exception.Response.StatusCode.value__
            if ($attempt -lt $MaxAttempts -and ($status -eq 429 -or $status -ge 500)) {
                $wait = [math]::Pow(2, $attempt)
                Write-Log "Graph returned $status, retrying in $wait s ($attempt/$MaxAttempts)..." 'WARN'
                Start-Sleep -Seconds $wait
            }
            else { throw }
        }
    }
}

function Save-DriveFolder {
    param([string]$DriveId, [string]$ItemId, [string]$LocalPath)

    if (-not (Test-Path -LiteralPath $LocalPath)) {
        New-Item -ItemType Directory -Path $LocalPath -Force | Out-Null
    }

    $uri = "v1.0/drives/$DriveId/items/$ItemId/children?`$top=999&`$select=id,name,size,file,folder,package"
    while ($uri) {
        $page = Invoke-GraphWithRetry -Uri $uri
        foreach ($item in $page.value) {
            $localItem = Join-Path $LocalPath (Get-SafeName $item.name)

            if ($item.folder) {
                Save-DriveFolder -DriveId $DriveId -ItemId $item.id -LocalPath $localItem
            }
            elseif ($item.file) {
                if ((Test-Path -LiteralPath $localItem) -and
                    (Get-Item -LiteralPath $localItem).Length -eq $item.size) {
                    $script:Skipped++
                    continue
                }
                try {
                    Invoke-GraphWithRetry -Uri "v1.0/drives/$DriveId/items/$($item.id)/content" -OutputFilePath $localItem | Out-Null
                    $script:Downloaded++
                    $script:Bytes += [int64]$item.size
                    Write-Log "Downloaded: $localItem"
                }
                catch {
                    $script:Failed.Add([pscustomobject]@{ Path = $localItem; Error = $_.Exception.Message })
                    Write-Log "FAILED: $localItem - $($_.Exception.Message)" 'ERROR'
                }
            }
            elseif ($item.package) {
                # OneNote notebooks etc. can't be downloaded as a single file.
                $script:Failed.Add([pscustomobject]@{ Path = $localItem; Error = "Skipped package item ($($item.package.type)) - export manually" })
                Write-Log "Skipped package ($($item.package.type)): $localItem" 'WARN'
            }
        }
        $uri = $page.'@odata.nextLink'
    }
}

#endregion

#region 1. Prompt for AD user ------------------------------------------------

try { Import-Module ActiveDirectory -ErrorAction Stop }
catch { throw 'The ActiveDirectory module (RSAT) is required. Install RSAT: Active Directory tools and try again.' }

do {
    $samAccountName = (Read-Host 'Enter the AD username (sAMAccountName) whose OneDrive you want to download').Trim()
} while ([string]::IsNullOrWhiteSpace($samAccountName))

try {
    $adUser = Get-ADUser -Identity $samAccountName -Properties UserPrincipalName, DisplayName, mail, Enabled
}
catch {
    throw "User '$samAccountName' was not found in Active Directory."
}

$upn = $adUser.UserPrincipalName
Write-Host ''
Write-Host "Found AD user:" -ForegroundColor Cyan
Write-Host "  Name    : $($adUser.DisplayName)"
Write-Host "  UPN     : $upn"
Write-Host "  Mail    : $($adUser.mail)"
Write-Host "  Enabled : $($adUser.Enabled)"
Write-Host ''

$confirm = Read-Host "Download this user's OneDrive to $DestinationRoot\$samAccountName ? (Y/N)"
if ($confirm -notmatch '^(y|yes)$') { Write-Host 'Cancelled.'; return }

#endregion

#region 2. Prepare destination + logging -------------------------------------

$userFolder = Join-Path $DestinationRoot (Get-SafeName $samAccountName)
New-Item -ItemType Directory -Path $userFolder -Force | Out-Null

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$script:LogFile = Join-Path $DestinationRoot "${samAccountName}_${stamp}.log"
$failFile       = Join-Path $DestinationRoot "${samAccountName}_${stamp}_failures.csv"

$script:Downloaded = 0
$script:Skipped    = 0
$script:Bytes      = [int64]0
$script:Failed     = [System.Collections.Generic.List[object]]::new()

Write-Log "Starting OneDrive archive for $samAccountName ($upn) -> $userFolder"

#endregion

#region 3. Connect to Microsoft Graph ----------------------------------------

try { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop }
catch { throw 'Microsoft.Graph.Authentication module is required: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser' }

if ($TenantId -and $ClientId -and $CertificateThumbprint) {
    Write-Log 'Connecting to Microsoft Graph (app-only, certificate)...'
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
}
else {
    Write-Log 'Connecting to Microsoft Graph (interactive sign-in)...'
    $connectParams = @{ Scopes = 'Files.Read.All', 'User.Read.All'; NoWelcome = $true }
    if ($TenantId) { $connectParams.TenantId = $TenantId }
    Connect-MgGraph @connectParams
}

#endregion

#region 4. Locate the user's OneDrive and download ---------------------------

try {
    $drive = Invoke-GraphWithRetry -Uri "v1.0/users/$upn/drive?`$select=id,webUrl,quota"
}
catch {
    Write-Log "Could not access the OneDrive for $upn. The user may not have a OneDrive provisioned, or you lack access. $($_.Exception.Message)" 'ERROR'
    Disconnect-MgGraph | Out-Null
    return
}

Write-Log "OneDrive: $($drive.webUrl)"
if ($drive.quota.used) {
    Write-Log ('OneDrive size (approx): {0:N2} GB' -f ($drive.quota.used / 1GB))
}

$started = Get-Date
try {
    Save-DriveFolder -DriveId $drive.id -ItemId 'root' -LocalPath $userFolder
}
finally {
    Disconnect-MgGraph | Out-Null
}

#endregion

#region 5. Summary -----------------------------------------------------------

$elapsed = (Get-Date) - $started
Write-Log '------------------------------------------------------------'
Write-Log "Finished in $($elapsed.ToString('hh\:mm\:ss'))"
Write-Log "Downloaded : $($script:Downloaded) file(s), $('{0:N2}' -f ($script:Bytes / 1MB)) MB"
Write-Log "Skipped    : $($script:Skipped) file(s) already present"
Write-Log "Failed     : $($script:Failed.Count) item(s)" $(if ($script:Failed.Count) { 'WARN' } else { 'INFO' })
Write-Log "Archive    : $userFolder"
Write-Log "Log        : $($script:LogFile)"

if ($script:Failed.Count) {
    $script:Failed | Export-Csv -Path $failFile -NoTypeInformation
    Write-Log "Failures   : $failFile" 'WARN'
}

#endregion
