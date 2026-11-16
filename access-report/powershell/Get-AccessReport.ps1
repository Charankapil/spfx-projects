<#
.SYNOPSIS
    Access Report from PowerShell: every site, subsite and library that chosen people and
    groups can reach, saved as the same CSV the web part and the monthly flow write.

.DESCRIPTION
    Makes the same SharePoint REST calls, in the same order, as the Access Report web part
    (services/AccessScan.ts and services/Sites.ts) and writes the same file:
    access-report-yyyy-MM-dd-HHmm.csv (UTC), and access-report-...-unread.csv for anything
    that could not be read.

    By default it reads who to report on, the options and the report folder from the site
    that hosts the Access Report web part (the same lists the web part and the flow use) and
    uploads the report there. Parameters override any of it, so it also works without the
    web part.

    Requirements:
      - PowerShell 7.2 or later, and the PnP.PowerShell module (Install-Module PnP.PowerShell).
      - An Entra ID app registration for PnP.PowerShell, which is needed to sign in at all since
        September 2024 (create one with Register-PnPEntraIDAppForInteractiveLogin; an admin
        consents once). The web part and the flow need no such registration - this is the one
        extra requirement of the PowerShell route.
      - The account that signs in sees what it can open: for a tenant-wide report, a site
        collection administrator of the sites to cover.

.PARAMETER HomeSiteUrl
    The site that hosts the Access Report web part: its lists hold the people and groups,
    the options and the report folder. Also where -SiteUrl and the search run from.

.PARAMETER ClientId
    Client id of your PnP.PowerShell app registration. Leave it out if you already ran
    Connect-PnPOnline to HomeSiteUrl in this session.

.PARAMETER Subject
    Who to report on, instead of the Subjects list: e-mail addresses (people) or SharePoint
    claims (c:0t.c|tenant|<group id>, c:0o.c|federateddirectoryclaimprovider|<group id>, ...).
    Use "Label=login" to choose the name shown in the report.

.PARAMETER SiteUrl
    Site collections to cover, instead of every site the account can find through search.

.PARAMETER SiteFilter
    Only site collections whose title or address contains this text. Overrides the setting.

.PARAMETER IncludeLibraries
    $true or $false. Overrides the setting (default: the setting, or $true).

.PARAMETER ReportFolder
    Server-relative folder in HomeSiteUrl to upload the report to. Overrides the setting.

.PARAMETER OutputPath
    Local folder for the CSV files. Default: the current folder.

.PARAMETER NoUpload
    Only write the local files.

.EXAMPLE
    ./Get-AccessReport.ps1 -HomeSiteUrl https://contoso.sharepoint.com/sites/it -ClientId 00000000-0000-0000-0000-000000000000

    The same report the monthly flow makes, uploaded to the report folder and saved locally.

.EXAMPLE
    ./Get-AccessReport.ps1 -HomeSiteUrl https://contoso.sharepoint.com/sites/it -ClientId <id> `
        -Subject 'Ann=ann@contoso.com','Auditors=c:0t.c|tenant|11111111-1111-1111-1111-111111111111' `
        -SiteUrl https://contoso.sharepoint.com/sites/fin -IncludeLibraries $false -NoUpload

    One-off report for two subjects on one site collection (and its subsites), local file only.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$HomeSiteUrl,
    [string]$ClientId,
    [string[]]$Subject,
    [string[]]$SiteUrl,
    [string]$SiteFilter,
    [Nullable[bool]]$IncludeLibraries,
    [string]$ReportFolder,
    [string]$OutputPath = (Get-Location).Path,
    [switch]$NoUpload
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------
# Format: identical to services/reportCsv.ts

$ReportHeaders = @('Who', 'Site URL', 'Site title', 'Object', 'Name', 'URL', 'Access', 'How', 'Own permissions', 'Site collection')
$UnreadHeaders = @('Who', 'Site URL', 'Site title', 'Step', 'Status', 'Error')
$AllSubjects = '(all)'
$NoAccess = 'No access'
$SiteOrders = @('Created:ascending', 'LastModifiedTime:ascending', '')

function ConvertTo-ReportCell([object]$Value) {
    $s = if ($null -eq $Value) { '' } else { [string]$Value }
    # A title such as =HYPERLINK(...) must not run as a formula in Excel.
    if ($s -cmatch '^[=+\-@\t\r]') { $s = "'" + $s }
    return '"' + $s.Replace('"', '""') + '"'
}

function ConvertTo-ReportCsv([string[]]$Headers, [System.Collections.Generic.List[object]]$Rows) {
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append([char]0xFEFF)
    foreach ($r in @(, $Headers) + $Rows.ToArray()) {
        [void]$sb.Append((@($r | ForEach-Object { ConvertTo-ReportCell $_ }) -join ',')).Append("`r`n")
    }
    return $sb.ToString()
}

function Format-Login([string]$Value) {
    $v = $Value.Trim()
    if ($v.Contains('|')) {
        # A Microsoft 365 group is stored as its members claim; its owners are checked with it.
        if ($v.ToLowerInvariant().Contains('|federateddirectoryclaimprovider|')) { $v = $v -replace '_o$', '' }
        return $v
    }
    return 'i:0#.f|membership|' + $v.ToLowerInvariant()
}

function Test-Person([string]$Login) { return $Login.Trim() -match '^i:' }
function Test-M365([string]$Login) { return $Login.ToLowerInvariant().Contains('|federateddirectoryclaimprovider|') }

# services/permissions.ts permissionLevel()
function Get-PermissionLevel([double]$Low, [double]$High) {
    $has = { param($kind) ([math]::Floor($Low / [math]::Pow(2, $kind - 1)) % 2) -eq 1 }
    if ($Low -ge 4294967295 -and $High -ge 2147483647) { return 'Full control' }
    if ((& $has 19) -or (& $has 26)) { return 'Design / manage' }
    if ((& $has 2) -and (& $has 3) -and (& $has 4)) { return 'Edit / contribute' }
    if (& $has 1) { return 'Read' }
    return $NoAccess
}

function Get-Origin([string]$Url) { return ([regex]::Match($Url, '^(https://[^/]+)', 'IgnoreCase')).Groups[1].Value }
function Remove-TrailingSlash([string]$Url) { return $Url -replace '/+$', '' }
function ConvertTo-UriLiteral([string]$Value) { return "'" + [uri]::EscapeDataString($Value.Replace("'", "''")) + "'" }

# ---------------------------------------------------------------------------------------
# REST

class RestError : System.Exception {
    [string]$Status
    RestError([string]$status, [string]$message) : base($message) { $this.Status = $status }
}

function Get-StatusOf($ErrorRecord) {
    $e = $ErrorRecord.Exception
    foreach ($candidate in @($e, $e.InnerException)) {
        if ($null -eq $candidate) { continue }
        foreach ($name in 'StatusCode', 'Response') {
            $p = $candidate.PSObject.Properties[$name]
            if ($null -eq $p -or $null -eq $p.Value) { continue }
            $v = if ($name -eq 'Response') { Get-Prop $p.Value 'StatusCode' } else { $p.Value }
            if ($null -ne $v) { try { return [string][int]$v } catch { } }
        }
    }
    $m = [regex]::Match([string]$e.Message, '\b([45]\d\d)\b')
    if ($m.Success) { return $m.Groups[1].Value }
    return ''
}

function Get-MessageOf($ErrorRecord) {
    $m = [string]$ErrorRecord.Exception.Message
    # SharePoint's own message, when the error carries its JSON.
    $j = [regex]::Match($m, '"(?:message|value)"\s*:\s*"((?:[^"\\]|\\.)*)"')
    if ($j.Success) { $m = $j.Groups[1].Value }
    $m = ($m -replace '[\r\n]+', ' ').Trim()
    if ($m.Length -gt 300) { $m = $m.Substring(0, 300) }
    return $m
}

# One GET, with a few retries while SharePoint throttles (429 / 503) or a gateway times out.
function Invoke-Get([string]$Url) {
    for ($attempt = 0; ; $attempt++) {
        try {
            return Invoke-PnPSPRestMethod -Method Get -Url $Url
        }
        catch {
            $status = Get-StatusOf $_
            if ($attempt -lt 4 -and $status -in '429', '503', '502', '504') {
                Start-Sleep -Seconds ([math]::Min(120, 2 * [math]::Pow(2, $attempt)))
                continue
            }
            throw [RestError]::new($status, (Get-MessageOf $_))
        }
    }
}

# The rows of a collection response, whichever OData flavour it came in.
function Get-Rows($Json) {
    if ($null -eq $Json) { return @() }
    if ($Json -is [array]) { return @($Json) }
    foreach ($name in 'value', 'results') {
        $p = $Json.PSObject.Properties[$name]
        if ($null -ne $p) { return @($p.Value) }
    }
    $d = Get-Prop $Json 'd'
    if ($null -ne $d) { return @(ConvertTo-List $d) }
    return @()
}

# A collection value as an array: missing, a plain array, or { results: [...] } (verbose OData).
function ConvertTo-List($Value) {
    if ($null -eq $Value) { return , @() }
    $r = Get-Prop $Value 'results'
    if ($null -ne $r) { return , @($r) }
    return , @($Value)
}

function Get-Prop($Object, [string]$Name) {
    if ($null -eq $Object) { return $null }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

# ---------------------------------------------------------------------------------------
# Sites (services/Sites.ts)

function Get-SearchRows($Json) {
    $primary = Get-Prop $Json 'PrimaryQueryResult'
    if ($null -eq $primary) { $primary = Get-Prop (Get-Prop (Get-Prop $Json 'd') 'query') 'PrimaryQueryResult' }
    $table = Get-Prop (Get-Prop $primary 'RelevantResults') 'Table'
    foreach ($row in (ConvertTo-List (Get-Prop $table 'Rows'))) {
        if ($null -eq $row) { continue }
        $record = @{}
        foreach ($c in (ConvertTo-List (Get-Prop $row 'Cells'))) {
            if ($null -ne $c) { $v = Get-Prop $c 'Value'; $record[[string](Get-Prop $c 'Key')] = if ($null -eq $v) { '' } else { [string]$v } }
        }
        $record
    }
}

function Get-SitePage([string]$Web, [string]$Order, [int]$StartRow, [int]$RowLimit = 500, [string]$Select = 'Title,Path,WebTemplate') {
    $parts = @("querytext=$(ConvertTo-UriLiteral 'contentclass:STS_Site')", "selectproperties='$([uri]::EscapeDataString($Select))'", "rowlimit=$RowLimit")
    if ($StartRow) { $parts += "startrow=$StartRow" }
    $parts += 'trimduplicates=false'
    if ($Order) { $parts += "sortlist='$([uri]::EscapeDataString($Order))'" }
    return @(Get-SearchRows (Invoke-Get "$Web/_api/search/query?$($parts -join '&')"))
}

function Test-InScope([string]$Url, [string]$Title, [string]$Template, [string]$Origin, [string]$Filter) {
    if (-not $Url -or -not ($Url + '/').ToLowerInvariant().StartsWith($Origin.ToLowerInvariant() + '/')) { return $false }
    if (($Url + '/') -match '^https://[^/]+/personal/' -or $Template -match '^SPSPERS') { return $false }
    return ($Title + ' ' + $Url).ToLowerInvariant().Contains($Filter.Trim().ToLowerInvariant())
}

function Get-SiteCollections([string]$Web, [string]$Filter) {
    # Some tenants refuse some sort orders: the first one this tenant accepts is used.
    $order = $null
    $refusals = @()
    foreach ($o in $SiteOrders) {
        try { [void](Get-SitePage $Web $o 0 1 'Path'); $order = $o; break }
        catch [RestError] {
            if ($_.Exception.Status -ne '400') { throw }
            $refusals += "$(if ($o) { $o } else { 'no order' }): $($_.Exception.Message)"
        }
    }
    if ($null -eq $order) { throw "SharePoint search refused every way of listing the sites. $($refusals -join '; ')" }
    $origin = Get-Origin $Web
    $sites = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    for ($start = 0; $start -lt 50000; $start += 500) {
        $rows = Get-SitePage $Web $order $start
        foreach ($r in $rows) {
            $url = Remove-TrailingSlash ([string]$r['Path'])
            $title = if ($r['Title']) { [string]$r['Title'] } else { $url }
            if ($seen.ContainsKey($url.ToLowerInvariant()) -or -not (Test-InScope $url $title ([string]$r['WebTemplate']) $origin $Filter)) { continue }
            $seen[$url.ToLowerInvariant()] = $true
            $sites.Add([pscustomobject]@{ Url = $url; Title = $title })
        }
        Write-Progress -Activity 'Listing sites' -Status "$($sites.Count) found"
        if ($rows.Count -lt 500) { Write-Progress -Activity 'Listing sites' -Completed; return $sites }
    }
    Write-Warning 'Search stopped at 50,000 results: the report covers the sites listed so far.'
    return $sites
}

# ---------------------------------------------------------------------------------------
# Scanning (services/AccessScan.ts)

$ListsQuery = '$select=Id,Title,BaseType,HasUniqueRoleAssignments,IsSystemList,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=Hidden eq false&$top=5000'
$AssignmentsQuery = '$select=PrincipalId,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind&$expand=Member,RoleDefinitionBindings&$top=5000'
$SiteUserQuery = '$select=Id,LoginName,IsSiteAdmin,Groups/Id,Groups/Title&$expand=Groups'
$WebsQuery = '$select=Title,Url,HasUniqueRoleAssignments'

function Get-Level($Json) {
    if ($null -ne (Get-Prop $Json 'd')) { $Json = $Json.d }
    $bp = Get-Prop $Json 'GetUserEffectivePermissions'
    if ($null -eq $bp) { $bp = $Json }
    $low = Get-Prop $bp 'Low'; $high = Get-Prop $bp 'High'
    return Get-PermissionLevel ([double]$(if ($low) { $low } else { 0 })) ([double]$(if ($high) { $high } else { 0 }))
}

# Grants that reach a principal: made to it, or to a SharePoint group it is in. Limited Access left out.
function Get-Grants($Assignments, [int]$PrincipalId, [int[]]$GroupIds) {
    foreach ($a in $Assignments) {
        $principal = [int](Get-Prop $a 'PrincipalId')
        if ($principal -ne $PrincipalId -and $GroupIds -notcontains $principal) { continue }
        $names = @((ConvertTo-List (Get-Prop $a 'RoleDefinitionBindings')) |
            Where-Object { $null -ne $_ -and (Get-Prop $_ 'RoleTypeKind') -ne 1 -and [string](Get-Prop $_ 'Name') -notmatch '^limited access' } |
            ForEach-Object { [string](Get-Prop $_ 'Name') })
        if (-not $names.Count) { continue }
        $member = Get-Prop $a 'Member'
        $how = if ([int](Get-Prop $member 'PrincipalType') -eq 8) { "SharePoint group: $(Get-Prop $member 'Title')" } else { 'Direct' }
        [pscustomobject]@{ Level = ($names -join ', '); How = $how }
    }
}

function Get-Webs($Site, [System.Collections.Generic.List[object]]$Unread) {
    $root = Remove-TrailingSlash $Site.Url
    $webs = [System.Collections.Generic.List[object]]::new()
    $webs.Add([pscustomobject]@{ Url = $root; Title = $Site.Title; Object = 'Site'; Unique = ''; Collection = $root })
    for ($i = 0; $i -lt $webs.Count; $i++) {
        $w = $webs[$i]
        try {
            foreach ($x in (Get-Rows (Invoke-Get "$($w.Url)/_api/web/webs?$WebsQuery"))) {
                $url = Remove-TrailingSlash ([string](Get-Prop $x 'Url'))
                if (-not $url -or @($webs | Where-Object { $_.Url -eq $url }).Count) { continue }
                $title = [string](Get-Prop $x 'Title'); if (-not $title) { $title = $url }
                $webs.Add([pscustomobject]@{ Url = $url; Title = $title; Object = 'Subsite'; Unique = $(if ((Get-Prop $x 'HasUniqueRoleAssignments') -eq $true) { 'Yes' } else { 'No' }); Collection = $root })
            }
        }
        catch [RestError] { $Unread.Add(@($AllSubjects, $w.Url, $w.Title, 'Subsites', $_.Exception.Status, $_.Exception.Message)) }
    }
    return $webs
}

function Invoke-WebScan($W, $Subjects, [bool]$Libraries, $Rows, $Unread, [string]$Origin) {
    $web = $W.Url
    $addUnread = { param($who, $step, $err) $Unread.Add(@($who, $web, $W.Title, $step, $err.Status, $err.Message)) }
    $row = {
        param($who, $object, $name, $url, $level, $how, $unique)
        $self = $object -eq 'Site'
        $Rows.Add(@($who, $web, $W.Title, $(if ($self) { $W.Object } else { $object }), $name, $url, $level, $how, $(if ($self) { $W.Unique } else { $unique }), $W.Collection))
    }

    $inherited = @(); $own = @()
    if ($Libraries) {
        try {
            foreach ($l in (Get-Rows (Invoke-Get "$web/_api/web/lists?$ListsQuery"))) {
                if ((Get-Prop $l 'BaseType') -ne 1 -or (Get-Prop $l 'IsSystemList') -eq $true) { continue }
                $path = [string](Get-Prop (Get-Prop $l 'RootFolder') 'ServerRelativeUrl')
                $lib = [pscustomobject]@{ Id = [string](Get-Prop $l 'Id'); Title = [string](Get-Prop $l 'Title'); Url = $(if ($path) { $Origin + $path } else { '' }) }
                if ((Get-Prop $l 'HasUniqueRoleAssignments') -eq $true) { $own += $lib } else { $inherited += $lib }
            }
        }
        catch [RestError] { & $addUnread $AllSubjects 'Libraries' $_.Exception }
    }

    # People first, then groups - the order the web part and the flow work in.
    foreach ($s in @($Subjects | Where-Object { Test-Person $_.Login })) {
        $at = "@v=$(ConvertTo-UriLiteral $s.Login.Trim())"
        try { $siteLevel = Get-Level (Invoke-Get "$web/_api/web/getusereffectivepermissions(@v)?$at") }
        catch [RestError] { & $addUnread $s.Label 'Site' $_.Exception; continue }
        if ($siteLevel -ne $NoAccess) {
            & $row $s.Label 'Site' $W.Title $web $siteLevel 'Effective permission' ''
            foreach ($l in $inherited) { & $row $s.Label 'Library' $l.Title $l.Url $siteLevel 'Inherits from site' 'No' }
        }
        foreach ($l in $own) {
            try {
                $level = Get-Level (Invoke-Get "$web/_api/web/lists(guid'$($l.Id)')/getusereffectivepermissions(@v)?$at")
                if ($level -ne $NoAccess) { & $row $s.Label 'Library' $l.Title $l.Url $level 'Effective permission (own permissions)' 'Yes' }
            }
            catch [RestError] { & $addUnread $s.Label "Library: $($l.Title)" $_.Exception }
        }
    }

    foreach ($s in @($Subjects | Where-Object { -not (Test-Person $_.Login) })) {
        $login = $s.Login.Trim()
        $claims = if (Test-M365 $login) { @($login, $login + '_o') } else { @($login) }
        $filter = ($claims | ForEach-Object { "LoginName eq '$($_.Replace("'", "''"))'" }) -join ' or '
        try { $found = @(Get-Rows (Invoke-Get "$web/_api/web/siteusers?`$filter=$([uri]::EscapeDataString($filter))&$SiteUserQuery")) }
        catch [RestError] { & $addUnread $s.Label 'Site' $_.Exception; continue }
        foreach ($f in $found) {
            $who = if ((Test-M365 $login) -and ([string](Get-Prop $f 'LoginName')).ToLowerInvariant().EndsWith('_o')) { "$($s.Label) (owners)" } else { $s.Label }
            $id = [int](Get-Prop $f 'Id')
            $groupIds = [int[]]@((ConvertTo-List (Get-Prop $f 'Groups')) | Where-Object { $null -ne $_ } | ForEach-Object { [int](Get-Prop $_ 'Id') })
            if ((Get-Prop $f 'IsSiteAdmin') -eq $true) { & $row $who 'Site' $W.Title $web 'Full control' 'Site collection administrator' '' }
            try { $siteGrants = @(Get-Grants (Get-Rows (Invoke-Get "$web/_api/web/roleassignments?$AssignmentsQuery")) $id $groupIds) }
            catch [RestError] { & $addUnread $who 'Site permissions' $_.Exception; continue }
            foreach ($gr in $siteGrants) {
                & $row $who 'Site' $W.Title $web $gr.Level $gr.How ''
                foreach ($l in $inherited) { & $row $who 'Library' $l.Title $l.Url $gr.Level "Inherits from site - $($gr.How)" 'No' }
            }
            foreach ($l in $own) {
                try {
                    foreach ($gr in @(Get-Grants (Get-Rows (Invoke-Get "$web/_api/web/lists(guid'$($l.Id)')/roleassignments?$AssignmentsQuery")) $id $groupIds)) {
                        & $row $who 'Library' $l.Title $l.Url $gr.Level $gr.How 'Yes'
                    }
                }
                catch [RestError] { & $addUnread $who "Library: $($l.Title)" $_.Exception }
            }
        }
    }
}

# ---------------------------------------------------------------------------------------
# Configuration (services/Store.ts lists)

$HomeWeb = Remove-TrailingSlash $HomeSiteUrl.Trim()
$WebPath = $HomeWeb -replace '^https://[^/]+', ''

if ($ClientId) {
    Connect-PnPOnline -Url $HomeWeb -ClientId $ClientId -Interactive
}
else {
    try { [void](Get-PnPConnection) }
    catch { throw 'Not connected. Pass -ClientId (your PnP.PowerShell app registration), or run Connect-PnPOnline first.' }
}

function Get-ListItems([string]$Relative, [string]$Select) {
    $u = ConvertTo-UriLiteral "$WebPath/$Relative"
    return @(Get-Rows (Invoke-Get "$HomeWeb/_api/web/GetList(@u)/items?`$select=$Select&`$orderby=Id&`$top=5000&@u=$u"))
}

$settings = $null
$needLists = (-not $Subject) -or ($null -eq $IncludeLibraries) -or (-not $PSBoundParameters.ContainsKey('SiteFilter')) -or (-not $ReportFolder -and -not $NoUpload)
if ($needLists) {
    try { $settings = @(Get-ListItems 'Lists/AccessReportSettings' 'Id,ARIncludeLibraries,ARSiteFilter,ARReportFolder') | Select-Object -First 1 }
    catch [RestError] {
        if (-not $Subject) { throw "Could not read the Access Report lists in $HomeWeb ($($_.Exception.Message)). Set the site up from the web part, or pass -Subject." }
    }
}

$subjects = if ($Subject) {
    foreach ($x in $Subject) {
        $label, $login = if ($x -match '^([^=|]+)=(.+)$') { $Matches[1].Trim(), $Matches[2] } else { $x.Trim(), $x }
        [pscustomobject]@{ Label = $label; Login = (Format-Login $login) }
    }
}
else {
    Get-ListItems 'Lists/AccessReportSubjects' 'Id,Title,ARLogin,ARActive' |
        Where-Object { (Get-Prop $_ 'ARActive') -eq $true -and ([string](Get-Prop $_ 'ARLogin')).Trim() } |
        ForEach-Object { [pscustomobject]@{ Label = [string](Get-Prop $_ 'Title'); Login = ([string]$_.ARLogin).Trim() } }
}
$subjects = @($subjects)
if (-not $subjects.Count) { throw 'Nobody to report on: add people or groups in the web part, or pass -Subject.' }

$libraries = if ($null -ne $IncludeLibraries) { [bool]$IncludeLibraries } else { (Get-Prop $settings 'ARIncludeLibraries') -ne $false }
$filter = if ($PSBoundParameters.ContainsKey('SiteFilter')) { $SiteFilter } else { [string](Get-Prop $settings 'ARSiteFilter') }
$folder = if ($ReportFolder) { $ReportFolder } elseif ((Get-Prop $settings 'ARReportFolder')) { [string]$settings.ARReportFolder } else { "$WebPath/AccessReports" }
$folder = Remove-TrailingSlash $folder.Trim()

# ---------------------------------------------------------------------------------------
# Run

$origin = Get-Origin $HomeWeb
$sites = if ($SiteUrl) {
    @($SiteUrl | ForEach-Object { $u = Remove-TrailingSlash $_.Trim(); [pscustomobject]@{ Url = $u; Title = $u } })
}
else {
    @(Get-SiteCollections $HomeWeb $filter)
}

Write-Host "Checking $($sites.Count) site collections (and their subsites) for $($subjects.Count) people and groups$(if ($libraries) { ', with libraries' })."
$rows = [System.Collections.Generic.List[object]]::new()
$unread = [System.Collections.Generic.List[object]]::new()
$n = 0
foreach ($site in $sites) {
    $n++
    Write-Progress -Activity 'Access report' -Status "$n of $($sites.Count): $($site.Url)" -PercentComplete (100 * ($n - 1) / [math]::Max(1, $sites.Count))
    foreach ($w in (Get-Webs $site $unread)) {
        Invoke-WebScan $w $subjects $libraries $rows $unread $origin
    }
}
Write-Progress -Activity 'Access report' -Completed

$stamp = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd-HHmm', [cultureinfo]::InvariantCulture)
$utf8 = [System.Text.UTF8Encoding]::new($false)
[void](New-Item -ItemType Directory -Force -Path $OutputPath)
$reportName = "access-report-$stamp.csv"
$reportPath = Join-Path $OutputPath $reportName
[System.IO.File]::WriteAllText($reportPath, (ConvertTo-ReportCsv $ReportHeaders $rows), $utf8)
$written = @($reportPath)
if ($unread.Count) {
    $unreadPath = Join-Path $OutputPath "access-report-$stamp-unread.csv"
    [System.IO.File]::WriteAllText($unreadPath, (ConvertTo-ReportCsv $UnreadHeaders $unread), $utf8)
    $written += $unreadPath
}

if (-not $NoUpload) {
    if (-not ($folder + '/').ToLowerInvariant().StartsWith($WebPath.ToLowerInvariant() + '/')) {
        Write-Warning "$folder is not in $HomeWeb, so the report was not uploaded. The local files are in $OutputPath."
    }
    else {
        $siteRelative = $folder.Substring($WebPath.Length).TrimStart('/')
        foreach ($p in $written) { [void](Add-PnPFile -Path $p -Folder $siteRelative) }
        Write-Host "Uploaded to $folder."
    }
}

Write-Host "$($rows.Count) rows, $($unread.Count) not read."
$written | ForEach-Object { Get-Item $_ }
