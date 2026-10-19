<#
.SYNOPSIS
  Sets up ECM Foundation (module 1): the ECM tenant theme, the ECM Navigation
  list that drives the ECM Hub Navigator, and optionally the app package.

.DESCRIPTION
  Every step is idempotent: re-running the script updates the theme, adds
  missing list columns and only adds navigation rows that are not there yet
  (matched on Title + type + parent hub), so governance edits are never
  overwritten.

  Requires PnP.PowerShell 2.x and an Entra ID app registration for it
  (PnP no longer ships a shared one): pass its client id with -ClientId.
  Theme registration and app deployment need SharePoint administrator rights.

.EXAMPLE
  ./Deploy-EcmFoundation.ps1 `
    -TenantAdminUrl https://contoso-admin.sharepoint.com `
    -EcmHubUrl https://contoso.sharepoint.com/sites/ecm-hub `
    -ClientId 00000000-0000-0000-0000-000000000000 `
    -ApplyThemeTo https://contoso.sharepoint.com/sites/real-estate, https://contoso.sharepoint.com/sites/finance-hub `
    -PackagePath ../releases/ecm-foundation.sppkg
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory)] [string] $TenantAdminUrl,
  [Parameter(Mandatory)] [string] $EcmHubUrl,
  [Parameter(Mandatory)] [string] $ClientId,
  [string] $ListTitle = 'ECM Navigation',
  [string] $SeedFile = (Join-Path $PSScriptRoot 'navigation-seed.json'),
  [string] $ThemeFile = (Join-Path $PSScriptRoot 'theme/ecm-theme.json'),
  # Sites to apply the ECM theme to (hubs apply it to their associated sites' header only, so list each hub).
  [string[]] $ApplyThemeTo = @(),
  # Path to ecm-foundation.sppkg to upload and deploy tenant-wide. Omit to deploy it yourself.
  [string] $PackagePath,
  # Grants "Everyone except external users" read on the navigation list. The navigator
  # cannot load for anyone who cannot read the list, so do this unless the hub site's
  # visitors group already covers every employee.
  [switch] $GrantEveryoneRead,
  [switch] $SkipSeed
)

$ErrorActionPreference = 'Stop'
$origin = ([Uri]$EcmHubUrl).GetLeftPart([UriPartial]::Authority)
$hubServerRelative = ([Uri]$EcmHubUrl).AbsolutePath.TrimEnd('/')

function Connect-Site([string] $Url) {
  Connect-PnPOnline -Url $Url -ClientId $ClientId -Interactive
}

# --- 1. Tenant theme ---------------------------------------------------------
$theme = Get-Content $ThemeFile -Raw | ConvertFrom-Json
$palette = @{}
$theme.palette.PSObject.Properties | ForEach-Object { $palette[$_.Name] = $_.Value }

Write-Host "Registering tenant theme '$($theme.name)'" -ForegroundColor Cyan
Connect-Site $TenantAdminUrl
Add-PnPTenantTheme -Identity $theme.name -Palette $palette -IsInverted $theme.isInverted -Overwrite

# --- 2. App package ----------------------------------------------------------
if ($PackagePath) {
  Write-Host "Deploying $PackagePath to the tenant app catalog" -ForegroundColor Cyan
  $app = Add-PnPApp -Path $PackagePath -Scope Tenant -Overwrite -Publish -SkipFeatureDeployment

  # Deploying registers the navigator in the "Tenant Wide Extensions" list with the
  # placeholder properties from ClientSideInstance.xml; point it at the real hub.
  $catalogUrl = Get-PnPTenantAppCatalogUrl
  Connect-Site $catalogUrl
  $props = @{ configSiteUrl = $hubServerRelative; listTitle = $ListTitle } | ConvertTo-Json -Compress
  $entries = Get-PnPListItem -List 'Tenant Wide Extensions' |
    Where-Object { $_['TenantWideExtensionComponentId'] -eq '293dd12c-1019-42fd-b262-79b6f949e9b7' }
  foreach ($entry in $entries) {
    Set-PnPListItem -List 'Tenant Wide Extensions' -Identity $entry.Id -Values @{ TenantWideExtensionComponentProperties = $props } | Out-Null
  }
  if (-not $entries) {
    Write-Warning 'No Tenant Wide Extensions entry found yet. It appears a few minutes after deployment; re-run with -PackagePath to update it.'
  }
}

# --- 3. Navigation list ------------------------------------------------------
Write-Host "Preparing list '$ListTitle' on $EcmHubUrl" -ForegroundColor Cyan
Connect-Site $EcmHubUrl
$list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if (-not $list) {
  $list = New-PnPList -Title $ListTitle -Template GenericList -OnQuickLaunch:$false
}

$fields = @(
  '<Field Type="Choice" Name="EcmNavType" StaticName="EcmNavType" DisplayName="Type" Required="TRUE" Format="Dropdown"><CHOICES><CHOICE>Hub</CHOICE><CHOICE>Department site</CHOICE><CHOICE>Shared service</CHOICE><CHOICE>Hub resource</CHOICE><CHOICE>Header link</CHOICE></CHOICES></Field>',
  '<Field Type="Text" Name="EcmParentHub" StaticName="EcmParentHub" DisplayName="Parent hub" Description="Title of the hub this entry belongs to (department sites and hub resources)." />',
  '<Field Type="URL" Name="EcmUrl" StaticName="EcmUrl" DisplayName="URL" Format="Hyperlink" Required="TRUE" />',
  '<Field Type="Text" Name="EcmDescription" StaticName="EcmDescription" DisplayName="Description" Description="Shown under the site name; for hubs, the governance note." />',
  '<Field Type="Number" Name="EcmOrder" StaticName="EcmOrder" DisplayName="Order" Decimals="0"><Default>100</Default></Field>',
  '<Field Type="Boolean" Name="EcmRestricted" StaticName="EcmRestricted" DisplayName="Restricted" Description="Shown only to people who can open the site."><Default>0</Default></Field>',
  '<Field Type="Boolean" Name="EcmActive" StaticName="EcmActive" DisplayName="Active"><Default>1</Default></Field>',
  '<Field Type="Text" Name="EcmColor" StaticName="EcmColor" DisplayName="Hub colour" Description="Optional hex colour for a hub, e.g. #B5470E." />'
)
$existing = (Get-PnPField -List $ListTitle).InternalName
foreach ($xml in $fields) {
  $name = ([xml]$xml).Field.Name
  if ($existing -notcontains $name) {
    Add-PnPFieldFromXml -List $ListTitle -FieldXml $xml | Out-Null
    Write-Host "  added column $name"
  }
}

$view = Get-PnPView -List $ListTitle -Identity 'All Items'
Set-PnPView -List $ListTitle -Identity $view.Id -Fields 'Title', 'EcmNavType', 'EcmParentHub', 'EcmUrl', 'EcmOrder', 'EcmRestricted', 'EcmActive' | Out-Null

if ($GrantEveryoneRead) {
  Set-PnPList -Identity $ListTitle -BreakRoleInheritance -CopyRoleAssignments | Out-Null
  Set-PnPListPermission -Identity $ListTitle -User 'c:0-.f|rolemanager|spo-grid-all-users/' -AddRole 'Read'
  Write-Host '  granted Everyone except external users read access'
}

# --- 4. Seed rows ------------------------------------------------------------
if (-not $SkipSeed) {
  $seed = (Get-Content $SeedFile -Raw | ConvertFrom-Json).items
  $current = Get-PnPListItem -List $ListTitle -PageSize 500 | ForEach-Object {
    "$($_['Title'])|$($_['EcmNavType'])|$($_['EcmParentHub'])".ToLowerInvariant()
  }
  $added = 0
  foreach ($row in $seed) {
    $key = "$($row.Title)|$($row.EcmNavType)|$($row.EcmParentHub)".ToLowerInvariant()
    if ($current -contains $key) { continue }
    $url = if ($row.EcmUrl.StartsWith('/')) { "$origin$($row.EcmUrl)" } else { $row.EcmUrl }
    Add-PnPListItem -List $ListTitle -Values @{
      Title          = $row.Title
      EcmNavType     = $row.EcmNavType
      EcmParentHub   = $row.EcmParentHub
      EcmUrl         = "$url, $($row.Title)"
      EcmDescription = $row.EcmDescription
      EcmOrder       = $row.EcmOrder
      EcmRestricted  = [bool]$row.EcmRestricted
      EcmActive      = $true
    } | Out-Null
    $added++
  }
  Write-Host "  added $added navigation rows ($($seed.Count - $added) already present)"
}

# --- 5. Apply the theme ------------------------------------------------------
foreach ($site in $ApplyThemeTo) {
  Write-Host "Applying '$($theme.name)' to $site" -ForegroundColor Cyan
  Connect-Site $site
  Set-PnPWebTheme -Theme $theme.name
}

Write-Host 'Done.' -ForegroundColor Green
