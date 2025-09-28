<# 
.SYNOPSIS
  Detects Microsoft 365 Groups with Visibility = Public (via Microsoft Graph) and remediates them to Private
  by applying a Sensitivity Label to the associated Site/Group. Logs actions, uploads logs to SharePoint,
  and relies on a Power Automate flow that triggers when files named "log*.txt" are added.

.DESCRIPTION
  - Runs daily from Azure Automation.
  - Authenticates using App Registration (certificate or client secret).
  - Uses Microsoft Graph to enumerate all M365 Groups (Unified).
  - Filters groups where Visibility = "Public".
  - For each public group, resolves the associated Site URL / Group Id and applies the configured
    Sensitivity Label (configured to enforce Private). Uses PnP for the label operation & SharePoint upload.
  - Writes action and report logs locally, then uploads to a SharePoint library.
  - Your Power Automate flow: Trigger when filename starts with "log" and ends with ".txt";
    Parse the two lines that record "Public → Private" and post to Teams.

.NOTES
  Replace placeholders with your tenant values and real Sensitivity Label Id.
#>

param(
  # ===== Identity / Auth =====
  [Parameter(Mandatory=$true)][string]$TenantId,               # 00000000-0000-0000-0000-000000000000
  [Parameter(Mandatory=$true)][string]$ClientId,               # App Registration ID
  [Parameter(Mandatory=$true)][ValidateSet("Certificate","Secret")][string]$AuthMode,
  [string]$CertificatePath = "C:\certs\pubgrp-app.pfx",
  [string]$CertificatePassword = "<PFX-PASSWORD>",
  [string]$ClientSecret = "<APP-CLIENT-SECRET>",

  # ===== SharePoint locations for log upload =====
  [Parameter(Mandatory=$true)][string]$SPOAdminUrl,            # https://contoso-admin.sharepoint.com
  [Parameter(Mandatory=$true)][string]$LogsSiteUrl,            # https://contoso.sharepoint.com/sites/Governance
  [Parameter(Mandatory=$true)][string]$LogsLibrary = "PublicGroupEnforcement",   # Document library

  # ===== Behavior =====
  [string]$LogsBasePath = "C:\Logs\PublicGroupFix",
  [string]$SensitivityLabelId = "00000000-0000-0000-0000-00000000ABCD"           # DUMMY: replace with real label GUID
)

# =========================
# Environment / Modules
# =========================
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
[System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

# Required modules:
#   Microsoft.Graph (v2+)  -> Install-Module Microsoft.Graph -Scope AllUsers
#   PnP.PowerShell         -> Install-Module PnP.PowerShell  -Scope AllUsers
Import-Module Microsoft.Graph -ErrorAction Stop
Import-Module PnP.PowerShell  -ErrorAction Stop

# =========================
# Logging
# =========================
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
New-Item -Path $LogsBasePath -ItemType Directory -Force | Out-Null

$ActionLog = Join-Path $LogsBasePath "logs_$stamp.txt"     # <— Flow watches: startswith("log") && endswith(".txt")
$ReportLog = Join-Path $LogsBasePath "report_$stamp.txt"

function Write-Log([string]$m){ ("{0}  {1}" -f (Get-Date -f "yyyy-MM-dd HH:mm:ss"), $m) | Tee-Object -FilePath $ActionLog -Append }
function Write-Report([string]$m){ $m | Tee-Object -FilePath $ReportLog -Append }

Write-Log "Job start"

# =========================
# AUTH: Graph + PnP
# =========================

function Connect-GraphApp {
  Write-Log "Connecting to Microsoft Graph (App)…"
  if ($AuthMode -eq "Certificate") {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertificatePath, $CertificatePassword, "Exportable,PersistKeySet")
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -Certificate $cert -NoWelcome
  } else {
    $secure = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -ClientSecret $secure -NoWelcome
  }
  Select-MgProfile -Name "beta" | Out-Null
}

function Connect-PnP($Url){
  Write-Log "Connecting PnP to $Url …"
  if ($AuthMode -eq "Certificate") {
    $secure = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
    Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $TenantId -CertificatePath $CertificatePath -CertificatePassword $secure -WarningAction SilentlyContinue
  } else {
    Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $TenantId -ClientSecret $ClientSecret -WarningAction SilentlyContinue
  }
}

Connect-GraphApp
Connect-PnP -Url $SPOAdminUrl

# =========================
# DISCOVER: All M365 Groups (Unified) via Graph
# =========================
Write-Log "Querying Microsoft 365 groups from Graph…"

# Get only Unified groups; pull key fields needed (id, displayName, visibility, resourceProvisioningOptions)
# Paging: iterate until complete
$PublicGroups = @()

$select = "id,displayName,visibility,resourceProvisioningOptions"
$filter = "groupTypes/any(a:a eq 'Unified')"

$response = Get-MgGroup -Filter $filter -CountVariable total -ConsistencyLevel eventual -All -Property $select -ErrorAction Stop
foreach($g in $response){
  if ($g.visibility -eq "Public") {
    $PublicGroups += $g
  }
}

Write-Log "Total Unified groups: $($response.Count); Public groups found: $($PublicGroups.Count)"

# =========================
# REPORT public groups (before remediation)
# =========================
if ($PublicGroups.Count -gt 0) {
  Write-Report "Public groups identified (pre-fix):"
  foreach($g in $PublicGroups){
    $line = "{0}  |  {1}" -f $g.Id, $g.DisplayName
    Write-Report $line
  }
} else {
  Write-Report "No public groups found."
}

# =========================
# REMEDIATE: Make Private using Sensitivity Label
# =========================
foreach($g in $PublicGroups){

  Write-Log "Processing group: $($g.DisplayName) ($($g.Id))"

  # Attempt to resolve the site URL (try PnP helper first)
  try {
    # PnP cmdlet can return SharePoint site URL if group is connected
    $pg = Get-PnPMicrosoft365Group -Identity $g.Id -IncludeSiteUrl -ErrorAction Stop
    $siteUrl = $pg.SiteUrl
  } catch { $siteUrl = $null }

  if (-not $siteUrl) {
    # Fallback via Graph: /groups/{id}/drive or /sites/root should contain the hostname;
    # For simplicity we continue even if no siteUrl (Set-PnP* can also target group itself).
    Write-Log "SiteUrl not resolved via PnP; proceeding with group-level operation."
  }

  # Apply sensitivity label → designed to enforce Private visibility
  try {
    if ($siteUrl) {
      Connect-PnP -Url $siteUrl
      # Option 1: Set site label (supported in modern tenants)
      Set-PnPSite -SensitivityLabel $SensitivityLabelId -ErrorAction Stop
    } else {
      # Option 2: Try group-level (less consistent across tenants)
      Set-PnPMicrosoft365Group -Identity $g.Id -SensitivityLabel $SensitivityLabelId -ErrorAction Stop
    }

    $msg = "Remediated group from Public → Private using sensitivity label. GroupId=$($g.Id) Name='$($g.DisplayName)'"
    Write-Log $msg
    # Log (for Flow to extract 2 lines)
    Write-Log "ACTION: PUBLIC→PRIVATE  GroupId=$($g.Id)  Name='$($g.DisplayName)'"
    Write-Log "LABEL:  $SensitivityLabelId"
  }
  catch {
    Write-Log "ERROR: Failed to apply label to $($g.Id) :: $_"
  }
}

# =========================
# UPLOAD LOGS to SharePoint (Power Automate watches for 'log*.txt')
# =========================
try {
  Connect-PnP -Url $LogsSiteUrl
  $folderServerRel = (Get-PnPList -Identity $LogsLibrary).RootFolder.ServerRelativeUrl

  foreach($f in @($ActionLog, $ReportLog)){
    if (Test-Path $f){
      Add-PnPFile -Path $f -Folder $folderServerRel -Values @{ CheckInComment = "Daily public→private enforcement" } | Out-Null
      Write-Log "Uploaded $(Split-Path $f -Leaf) to $folderServerRel"
    }
  }
}
catch {
  Write-Log "ERROR: Failed to upload logs to SharePoint :: $_"
}

Write-Log "Job end"
