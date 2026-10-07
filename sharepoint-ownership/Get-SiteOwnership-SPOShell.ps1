<#
.SYNOPSIS
  Option 1: owners and members of every site collection using ONLY the SharePoint Online
  Management Shell. No app registration, no PnP, no Graph.

.DESCRIPTION
  Owner = member of the site's Owners group (any SharePoint group that has Full Control).
  Site collection admins are NOT reported as owners.
  Member = member of the site's Members group (any group with Edit/Contribute).
  For group-connected sites, the M365 group appears as a claim (c:0o.c|federateddirectoryclaimprovider|<id>
  for members, ..._o for owners). Add -ExpandM365Groups to expand them with Exchange Online.

  LIMITATION: the SPO module works at site collection level only. Subsites with unique
  permissions are not covered. Use Get-SiteOwnership.ps1 for subsites.

.NOTES
  Requires: Microsoft.Online.SharePoint.PowerShell (Windows PowerShell 5.1 or PowerShell 7 with -UseWindowsPowerShell).
  Optional: ExchangeOnlineManagement for -ExpandM365Groups, ImportExcel for -ExportExcel.
  Not tested against a live tenant from the authoring environment. Try -SiteUrlFilter on one site first.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantName,         # "contoso"
    [string]$OutputFolder = (Get-Location).Path,
    [string]$SiteUrlFilter = '*',
    [switch]$GrantTemporaryAccess,                      # add yourself as site admin while reading, then remove (audited)
    [string]$MyLogin,                                   # your UPN, required with -GrantTemporaryAccess
    [switch]$ExpandM365Groups,
    [switch]$ExportExcel
)

$ErrorActionPreference = 'Stop'
$stamp = Get-Date -Format 'yyyyMMdd_HHmm'
$OutFile  = Join-Path $OutputFolder "SiteOwnership_SPO_$stamp.csv"
$ErrFile  = Join-Path $OutputFolder "SiteOwnership_SPO_Errors_$stamp.csv"
$XlsxFile = Join-Path $OutputFolder "SiteOwnership_SPO_$stamp.xlsx"

if ($ExportExcel -and -not (Get-Module -ListAvailable ImportExcel)) {
    throw 'ImportExcel module not found. Run: Install-Module ImportExcel -Scope CurrentUser'
}

$SkipTemplates = 'SRCHCEN#0','SPSMSITEHOST#0','APPCATALOG#0','POINTPUBLISHINGHUB#0',
                 'POINTPUBLISHINGTOPIC#0','EDISC#0','REDIRECTSITE#0','TEAMCHANNEL#0','RESEARCHCENTER#0'
$IgnoreLogin = '^(SHAREPOINT\\system|app@sharepoint|c:0\(\.s\|true|c:0-\.f\|rolemanager\|spo-grid-all-users)'

Write-Host "Connecting to https://$TenantName-admin.sharepoint.com ..."
Connect-SPOService -Url "https://$TenantName-admin.sharepoint.com"
if ($ExpandM365Groups) { Connect-ExchangeOnline -ShowBanner:$false }

if ($GrantTemporaryAccess -and -not $MyLogin) { throw '-GrantTemporaryAccess needs -MyLogin (your UPN, e.g. you@contoso.com)' }

$Rows = New-Object System.Collections.Generic.List[object]
$Errors = New-Object System.Collections.Generic.List[object]
$GroupCache = @{}

function Get-M365Links([string]$GroupId, [string]$Type) {
    $k = "$GroupId|$Type"
    if (-not $GroupCache.ContainsKey($k)) {
        $GroupCache[$k] = @(Get-UnifiedGroupLinks -Identity $GroupId -LinkType $Type -ResultSize Unlimited -ErrorAction SilentlyContinue |
            ForEach-Object { [pscustomobject]@{ Name = $_.DisplayName; Login = $_.WindowsLiveID; Email = $_.PrimarySmtpAddress } })
    }
    $GroupCache[$k]
}

function Add-Principal($site, [string]$role, [string]$groupTitle, [string]$login) {
    if ($login -match $IgnoreLogin) { return }
    if ($login -match 'federateddirectoryclaimprovider\|([0-9a-f-]{36})(_o)?$') {
        $gid = $Matches[1]; $isOwner = [bool]$Matches[2]
        if ($ExpandM365Groups) {
            foreach ($u in (Get-M365Links $gid $(if ($isOwner) { 'Owners' } else { 'Members' }))) {
                $Rows.Add([pscustomobject]@{ SiteUrl=$site.Url; SiteTitle=$site.Title; M365GroupId=$site.GroupId
                    Role = $(if ($isOwner) { 'Owner' } else { $role }); Name=$u.Name; Login=$u.Login; Email=$u.Email
                    PrincipalType='User'; Source="M365 group $(if ($isOwner) {'owner'} else {'member'}) via $groupTitle" })
            }
            return
        }
        $type = if ($isOwner) { 'M365GroupOwners (unexpanded)' } else { 'M365GroupMembers (unexpanded)' }
        $Rows.Add([pscustomobject]@{ SiteUrl=$site.Url; SiteTitle=$site.Title; M365GroupId=$site.GroupId
            Role = $(if ($isOwner) { 'Owner' } else { $role }); Name=$login; Login=$login; Email=''
            PrincipalType=$type; Source="SP group: $groupTitle" })
        return
    }
    $type = if ($login -match '#ext#') { 'Guest' } elseif ($login -match '^c:0t\.c\|tenant\|') { 'EntraSecurityGroup (unexpanded)' } elseif ($login -match '^i:0#\.f\|membership\|') { 'User' } else { 'Other' }
    $Rows.Add([pscustomobject]@{ SiteUrl=$site.Url; SiteTitle=$site.Title; M365GroupId=$site.GroupId
        Role=$role; Name=($login -replace '^.*\|',''); Login=$login; Email=$(if ($login -match '\|([^|]+@[^|]+)$') { $Matches[1] } else { '' })
        PrincipalType=$type; Source="SP group: $groupTitle" })
}

$sites = Get-SPOSite -Limit All -Detailed | Where-Object {
    $_.Template -notin $SkipTemplates -and $_.Url -like $SiteUrlFilter -and $_.Url -notmatch '-my\.sharepoint\.com' }
Write-Host "Processing $($sites.Count) site collections."

$i = 0
foreach ($s in $sites) {
    $i++; Write-Progress -Activity 'Site ownership' -Status $s.Url -PercentComplete (100 * $i / $sites.Count)
    $site = [pscustomobject]@{ Url = $s.Url; Title = $s.Title; GroupId = $(if ($s.GroupId -and $s.GroupId -ne [guid]::Empty) { "$($s.GroupId)" } else { '' }) }
    $granted = $false
    try {
        if ($GrantTemporaryAccess) {
            Set-SPOUser -Site $s.Url -LoginName $MyLogin -IsSiteCollectionAdmin $true | Out-Null
            $granted = $true
        }
        foreach ($g in (Get-SPOSiteGroup -Site $s.Url)) {
            $roles = @($g.Roles)
            $role = if ($roles -contains 'Full Control') { 'Owner' }
                    elseif ($roles -contains 'Edit' -or $roles -contains 'Contribute') { 'Member' }
                    else { $null }     # Read-only groups (visitors) are ignored
            if (-not $role) { continue }
            foreach ($login in @($g.Users)) { Add-Principal $site $role $g.Title $login }
        }
    }
    catch { $Errors.Add([pscustomobject]@{ Url = $s.Url; Error = $_.Exception.Message }) }
    finally {
        if ($granted) {
            try { Set-SPOUser -Site $s.Url -LoginName $MyLogin -IsSiteCollectionAdmin $false | Out-Null }
            catch { $Errors.Add([pscustomobject]@{ Url = $s.Url; Error = "FAILED TO REMOVE TEMP ADMIN: $($_.Exception.Message)" }) }
        }
    }
}

# Drop duplicates, and drop Member rows for anyone who is also Owner on the same site
$ownerKeys = @{}
$Rows | Where-Object Role -eq 'Owner' | ForEach-Object { $ownerKeys["$($_.SiteUrl)|$($_.Login)"] = $true }
$final = $Rows | Where-Object { $_.Role -ne 'Member' -or -not $ownerKeys.ContainsKey("$($_.SiteUrl)|$($_.Login)") } |
    Sort-Object SiteUrl, Role, Login -Unique
$final | Export-Csv $OutFile -NoTypeInformation -Encoding UTF8

$withOwners = $final | Where-Object Role -eq 'Owner' | Select-Object -ExpandProperty SiteUrl -Unique
$noOwners = $sites | Where-Object { $_.Url -notin $withOwners } | Select-Object Url, Title, Template, GroupId
if ($Errors.Count) { $Errors | Export-Csv $ErrFile -NoTypeInformation -Encoding UTF8 }

if ($ExportExcel) {
    Import-Module ImportExcel
    $x = @{ Path = $XlsxFile; AutoSize = $true; FreezeTopRow = $true; BoldTopRow = $true; AutoFilter = $true }
    $final | Export-Excel @x -WorksheetName 'Detail' -ClearSheet
    if (@($noOwners).Count) { $noOwners | Export-Excel @x -WorksheetName 'No owners' -Append }
    if ($Errors.Count)      { $Errors   | Export-Excel @x -WorksheetName 'Errors' -Append }
    Write-Host "Excel workbook -> $XlsxFile"
}

Write-Host "Done. $($final.Count) rows -> $OutFile"
Write-Host "Sites without an owner found: $(@($noOwners).Count)"
if ($Errors.Count) { Write-Warning "$($Errors.Count) errors -> $ErrFile (access denied usually means -GrantTemporaryAccess is needed)" }
