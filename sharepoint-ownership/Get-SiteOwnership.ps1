<#
.SYNOPSIS
  Reports the real owners and members of every site collection and subsite,
  including the owners/members of the Microsoft 365 group behind group-connected sites.

.DESCRIPTION
  "Owner" here means a principal that SharePoint treats as an owner of the web:
    - member of the web's associated Owners group
    - member of any group (or direct principal) holding Full Control on the web
    - owner of the connected M365 group (the "_o" claim)
  Site collection administrators are deliberately NOT reported as owners. That list
  usually contains tenant admins and the primary-admin field is not an ownership signal.

  No app registration is needed. Sign-in uses a Microsoft first-party public client
  (SharePoint Online Management Shell) unless you pass your own -ClientId.
  M365 group expansion uses Exchange Online PowerShell (no Graph consent needed).

.NOTES
  Requires: PnP.PowerShell (v2+), ExchangeOnlineManagement.
  Optional: Microsoft.Graph.Groups, only with -ResolveSecurityGroups, to expand Entra security groups.
  Optional: ImportExcel, only with -ExportExcel. Excel does not need to be installed.
  Not tested against a live tenant from the authoring environment. Run with -SiteUrlFilter
  on one site first.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantName,        # "contoso" for contoso.sharepoint.com
    [string]$ClientId,                                 # optional own app; default is the SPO Management Shell first-party app
    [string]$OutputFolder = (Get-Location).Path,
    [string]$SiteUrlFilter = '*',                      # wildcard on site URL, e.g. '*/sites/hr*'
    [switch]$IncludeVisitors,
    [switch]$SkipSubsites,
    [switch]$GrantTemporaryAccess,                     # add yourself as site admin while reading, then remove (audited)
    [switch]$ResolveSecurityGroups,                    # expand Entra security groups via Microsoft Graph
    [switch]$ExportExcel                               # also write one .xlsx workbook (needs ImportExcel module, not Excel itself)
)

$ErrorActionPreference = 'Stop'
$SpoManagementShellClientId = '9bc3ab49-b65d-410a-85ad-de819febfddc'
$AdminUrl = "https://$TenantName-admin.sharepoint.com"
$stamp    = Get-Date -Format 'yyyyMMdd_HHmm'
$OutFile  = Join-Path $OutputFolder "SiteOwnership_$stamp.csv"
$NoOwnerFile = Join-Path $OutputFolder "SitesWithoutOwners_$stamp.csv"
$ErrFile  = Join-Path $OutputFolder "SiteOwnershipErrors_$stamp.csv"
$XlsxFile = Join-Path $OutputFolder "SiteOwnership_$stamp.xlsx"

# Templates that are not real collaboration sites
$SkipTemplates = 'SRCHCEN#0','SPSMSITEHOST#0','APPCATALOG#0','POINTPUBLISHINGHUB#0',
                 'POINTPUBLISHINGTOPIC#0','EDISC#0','REDIRECTSITE#0','TEAMCHANNEL#0','RESEARCHCENTER#0'
# Principals that are never human owners
$SystemLogin = '^(SHAREPOINT\\system|app@sharepoint|c:0\(\.s\|true|c:0-\.f\|rolemanager\|spo-grid-all-users|c:0t\.c\|tenant\|.*)$'
$SystemTitle = 'Company Administrator','SharePoint Service Administrator','System Account','SharePoint App'

if ($ExportExcel -and -not (Get-Module -ListAvailable ImportExcel)) {
    throw 'ImportExcel module not found. Run: Install-Module ImportExcel -Scope CurrentUser'
}

$Rows   = New-Object System.Collections.Generic.List[object]
$Errors = New-Object System.Collections.Generic.List[object]
$GroupCache = @{}

function Connect-Pnp([string]$Url) {
    $p = @{ Url = $Url; ReturnConnection = $true }
    if ($ClientId) { $p.ClientId = $ClientId; $p.Interactive = $true }
    else           { $p.ClientId = $SpoManagementShellClientId; $p.DeviceLogin = $true }
    Connect-PnPOnline @p
}

function Get-PrincipalKind($m) {
    $l = $m.LoginName
    if ($l -match 'federateddirectoryclaimprovider\|([0-9a-f-]{36})(_o)?$') {
        return @{ Kind = if ($Matches[2]) { 'M365GroupOwners' } else { 'M365GroupMembers' }; GroupId = $Matches[1] }
    }
    if ($l -match '^c:0t\.c\|tenant\|([0-9a-f-]{36})$' -and $m.Title -notin $SystemTitle) {
        return @{ Kind = 'EntraSecurityGroup'; GroupId = $Matches[1] }
    }
    if ($l -match '#ext#') { return @{ Kind = 'Guest' } }
    if ($l -match '^i:0#\.f\|membership\|') { return @{ Kind = 'User' } }
    if ($m.PrincipalType -eq 'SharePointGroup') { return @{ Kind = 'SharePointGroup' } }
    return @{ Kind = 'Other' }
}

function Get-M365GroupLinks([string]$GroupId, [string]$Type) {
    $key = "$GroupId|$Type"
    if (-not $GroupCache.ContainsKey($key)) {
        $GroupCache[$key] = @(Get-UnifiedGroupLinks -Identity $GroupId -LinkType $Type -ResultSize Unlimited -ErrorAction SilentlyContinue |
            ForEach-Object { [pscustomobject]@{ Name = $_.DisplayName; Login = $_.WindowsLiveID; Email = $_.PrimarySmtpAddress } })
    }
    $GroupCache[$key]
}

function Get-EntraGroupUsers([string]$GroupId) {
    $key = "sec|$GroupId"
    if (-not $GroupCache.ContainsKey($key)) {
        $GroupCache[$key] = @(Get-MgGroupTransitiveMember -GroupId $GroupId -All |
            Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user' } |
            ForEach-Object { [pscustomobject]@{ Name = $_.AdditionalProperties.displayName; Login = $_.AdditionalProperties.userPrincipalName; Email = $_.AdditionalProperties.mail } })
    }
    $GroupCache[$key]
}

# Expands one principal into flat user rows
function Expand-Principal($m, [string]$Source) {
    $k = Get-PrincipalKind $m
    switch ($k.Kind) {
        'M365GroupOwners'  { Get-M365GroupLinks $k.GroupId 'Owners'  | ForEach-Object { @{ N=$_.Name; L=$_.Login; E=$_.Email; T='User'; S="M365 group owner ($($k.GroupId))" } } }
        'M365GroupMembers' { Get-M365GroupLinks $k.GroupId 'Members' | ForEach-Object { @{ N=$_.Name; L=$_.Login; E=$_.Email; T='User'; S="M365 group member ($($k.GroupId))" } } }
        'EntraSecurityGroup' {
            if ($ResolveSecurityGroups) {
                Get-EntraGroupUsers $k.GroupId | ForEach-Object { @{ N=$_.Name; L=$_.Login; E=$_.Email; T='User'; S="Entra group: $($m.Title)" } }
            } else {
                @{ N=$m.Title; L=$m.LoginName; E=$m.Email; T='UnexpandedSecurityGroup'; S="$Source" }
            }
        }
        'Other' {
            if ($m.LoginName -match $SystemLogin -or $m.Title -in $SystemTitle) { return }
            @{ N=$m.Title; L=$m.LoginName; E=$m.Email; T='Other'; S=$Source }
        }
        default { @{ N=$m.Title; L=$m.LoginName; E=$m.Email; T=$k.Kind; S=$Source } }
    }
}

function Add-Rows($site, $web, [bool]$isRoot, [bool]$inherits, [string]$role, $principals, [string]$source) {
    foreach ($p in $principals) {
        foreach ($u in @(Expand-Principal $p $source)) {
            if (-not $u) { continue }
            $Rows.Add([pscustomobject]@{
                SiteUrl = $site.Url; SiteTitle = $site.Title; M365GroupId = $site.GroupId
                WebUrl = $web.Url; WebTitle = $web.Title; IsSubsite = -not $isRoot
                InheritsPermissions = $inherits
                Role = $role; Name = $u.N; Login = $u.L; Email = $u.E; PrincipalType = $u.T
                Source = if ($u.S -match '^(M365|Entra)') { $u.S } else { "$source" }
            })
        }
    }
}

function Get-GroupUsers($conn, $groupId) {
    Get-PnPGroupMember -Group $groupId -Connection $conn
}

function Read-Web($conn, $site, $webId, [bool]$isRoot) {
    $w = Get-PnPWeb -Identity $webId -Connection $conn `
        -Includes Title,Url,HasUniqueRoleAssignments,AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup,RoleAssignments
    $inherits = -not $w.HasUniqueRoleAssignments
    if ($isRoot) { $inherits = $false }

    # Owners: associated Owners group
    $ownerGroupIds = @()
    if ($w.AssociatedOwnerGroup.Id) {
        $ownerGroupIds += $w.AssociatedOwnerGroup.Id
        Add-Rows $site $w $isRoot $inherits 'Owner' (Get-GroupUsers $conn $w.AssociatedOwnerGroup.Id) "SP group: $($w.AssociatedOwnerGroup.Title)"
    }

    # Owners: any other principal with Full Control on this web
    foreach ($ra in $w.RoleAssignments) {
        Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings, Member -Connection $conn | Out-Null
        if (-not ($ra.RoleDefinitionBindings | Where-Object { $_.RoleTypeKind -eq 'Administrator' })) { continue }
        if ($ra.Member.Id -in $ownerGroupIds) { continue }
        if ($ra.Member.PrincipalType -eq 'SharePointGroup') {
            Add-Rows $site $w $isRoot $inherits 'Owner' (Get-GroupUsers $conn $ra.Member.Id) "Full Control via SP group: $($ra.Member.Title)"
        } else {
            Add-Rows $site $w $isRoot $inherits 'Owner' @($ra.Member) 'Direct Full Control'
        }
    }

    # Members / visitors
    if ($w.AssociatedMemberGroup.Id) {
        Add-Rows $site $w $isRoot $inherits 'Member' (Get-GroupUsers $conn $w.AssociatedMemberGroup.Id) "SP group: $($w.AssociatedMemberGroup.Title)"
    }
    if ($IncludeVisitors -and $w.AssociatedVisitorGroup.Id) {
        Add-Rows $site $w $isRoot $inherits 'Visitor' (Get-GroupUsers $conn $w.AssociatedVisitorGroup.Id) "SP group: $($w.AssociatedVisitorGroup.Title)"
    }
    $w
}

# ---- connect -------------------------------------------------------------------------
Write-Host "Connecting to $AdminUrl ..."
$adminConn = Connect-Pnp $AdminUrl
Write-Host 'Connecting to Exchange Online (for M365 group owners/members) ...'
Connect-ExchangeOnline -ShowBanner:$false
if ($ResolveSecurityGroups) { Connect-MgGraph -Scopes 'GroupMember.Read.All' -NoWelcome }
$me = (Get-PnPProperty -ClientObject (Get-PnPWeb -Connection $adminConn) -Property CurrentUser -Connection $adminConn).LoginName

$sites = Get-PnPTenantSite -Detailed -Connection $adminConn |
    Where-Object { $_.Template -notin $SkipTemplates -and $_.Url -like $SiteUrlFilter -and $_.Url -notmatch '-my\.sharepoint\.com' }
Write-Host "Processing $($sites.Count) site collections."

$i = 0
foreach ($s in $sites) {
    $i++; Write-Progress -Activity 'Site ownership' -Status $s.Url -PercentComplete (100 * $i / $sites.Count)
    $site = [pscustomobject]@{ Url = $s.Url; Title = $s.Title; GroupId = if ($s.GroupId -and $s.GroupId -ne [guid]::Empty) { $s.GroupId } else { '' } }
    $granted = $false
    try {
        if ($GrantTemporaryAccess) {
            Set-PnPTenantSite -Identity $s.Url -Owners $me -Connection $adminConn
            $granted = $true
        }
        $conn = Connect-Pnp $s.Url
        $root = Read-Web $conn $site (Get-PnPWeb -Connection $conn).Id $true

        # Group-connected sites: always include the M365 group owners from the source of truth
        if ($site.GroupId) {
            Add-Rows $site $root $true $false 'Owner' @([pscustomobject]@{ LoginName = "c:0o.c|federateddirectoryclaimprovider|$($site.GroupId)_o"; Title = ''; Email = ''; PrincipalType = 'User' }) 'M365 group owner'
        }
        if (-not $SkipSubsites) {
            foreach ($sub in (Get-PnPSubWeb -Recurse -Connection $conn)) {
                try { Read-Web $conn $site $sub.Id $false | Out-Null }
                catch { $Errors.Add([pscustomobject]@{ Url = $sub.Url; Error = $_.Exception.Message }) }
            }
        }
    }
    catch { $Errors.Add([pscustomobject]@{ Url = $s.Url; Error = $_.Exception.Message }) }
    finally {
        if ($granted) {
            try { Remove-PnPSiteCollectionAdmin -Owners $me -Connection (Connect-Pnp $s.Url) } catch {
                $Errors.Add([pscustomobject]@{ Url = $s.Url; Error = "FAILED TO REMOVE TEMP ADMIN: $($_.Exception.Message)" })
            }
        }
    }
}

# ---- output --------------------------------------------------------------------------
# Drop duplicates, and drop Member rows for anyone who is Owner of the same web
$ownerKeys = @{}
$Rows | Where-Object Role -eq 'Owner' | ForEach-Object { $ownerKeys["$($_.WebUrl)|$($_.Login)"] = $true }
$final = $Rows | Where-Object { $_.Role -ne 'Member' -or -not $ownerKeys.ContainsKey("$($_.WebUrl)|$($_.Login)") } |
    Sort-Object WebUrl, Role, Login -Unique
$final | Export-Csv $OutFile -NoTypeInformation -Encoding UTF8

$withOwners = $final | Where-Object { $_.Role -eq 'Owner' -and $_.PrincipalType -ne 'UnexpandedSecurityGroup' } | Select-Object -ExpandProperty SiteUrl -Unique
$noOwners = $sites | Where-Object { $_.Url -notin $withOwners } | Select-Object Url, Title, Template, GroupId
$noOwners | Export-Csv $NoOwnerFile -NoTypeInformation -Encoding UTF8
if ($Errors.Count) { $Errors | Export-Csv $ErrFile -NoTypeInformation -Encoding UTF8 }

if ($ExportExcel) {
    Import-Module ImportExcel
    $excel = @{ Path = $XlsxFile; AutoSize = $true; FreezeTopRow = $true; BoldTopRow = $true; AutoFilter = $true }
    # One line per web and role, with the people joined, for a quick read
    $summary = $final | Group-Object WebUrl, Role | ForEach-Object {
        $f = $_.Group[0]
        [pscustomobject]@{
            SiteUrl = $f.SiteUrl; SiteTitle = $f.SiteTitle; WebUrl = $f.WebUrl; IsSubsite = $f.IsSubsite
            InheritsPermissions = $f.InheritsPermissions; Role = $f.Role; Count = $_.Count
            People = ($_.Group | ForEach-Object { if ($_.Email) { "$($_.Name) <$($_.Email)>" } else { $_.Name } }) -join '; '
        }
    } | Sort-Object WebUrl, Role
    $summary | Export-Excel @excel -WorksheetName 'Summary' -ClearSheet
    $final   | Export-Excel @excel -WorksheetName 'Detail' -Append
    if (@($noOwners).Count) { $noOwners | Export-Excel @excel -WorksheetName 'No owners' -Append }
    if ($Errors.Count)      { $Errors   | Export-Excel @excel -WorksheetName 'Errors' -Append }
    Write-Host "Excel workbook -> $XlsxFile"
}

Write-Host "Done. $($final.Count) rows -> $OutFile"
Write-Host "Sites without a resolvable owner: $(@($noOwners).Count) -> $NoOwnerFile"
if ($Errors.Count) { Write-Warning "$($Errors.Count) errors -> $ErrFile (access denied usually means -GrantTemporaryAccess is needed)" }
