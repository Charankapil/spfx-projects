<#
.SYNOPSIS
    Access Report from PowerShell, with the SharePoint Online Management Shell: no app
    registration and no client id. Site collection level.

.DESCRIPTION
    Signs in with Connect-SPOService (Microsoft's own SharePoint admin module) and, for every
    site collection, reports where the chosen people and groups get in:

      - as a site collection administrator (Get-SPOUser ... IsSiteAdmin), and
      - through the site's SharePoint groups (Get-SPOSiteGroup: each group's members and
        permission levels) - including, for a person, groups that contain "Everyone" or
        "Everyone except external users".

    It writes the same CSV format as the web part and the monthly flow (same columns, same
    quoting, UTF-8 with BOM), as access-report-yyyy-MM-dd-HHmm.csv in UTC, and
    access-report-...-unread.csv for site collections it could not read.

    What the Management Shell cannot see, and so this script does not report:
      - subsites and libraries (the module has no cmdlets for webs or lists);
      - permissions granted directly to a person or group, not through a SharePoint group;
      - membership through Entra ID or Microsoft 365 groups for a person (report on the group
        itself instead: pass its claim as a -Subject).
    The web part and the monthly flow report all of these. Use this script for a quick
    site-collection-level answer without any app registration.

    Requirements: Windows PowerShell 5.1 or PowerShell 7 with the module
    Microsoft.Online.SharePoint.PowerShell (Install-Module Microsoft.Online.SharePoint.PowerShell),
    and the SharePoint Administrator role. A SharePoint admin can read every site collection's
    groups and administrators without being a member of the sites.

.PARAMETER AdminUrl
    The SharePoint admin center address, e.g. https://contoso-admin.sharepoint.com.
    Leave it out if you already ran Connect-SPOService in this session.

.PARAMETER Subject
    Who to report on: e-mail addresses (people) or SharePoint claims, e.g.
    c:0t.c|tenant|<security group object id> or
    c:0o.c|federateddirectoryclaimprovider|<Microsoft 365 group object id>.
    Use "Label=login" to choose the name shown in the report.

.PARAMETER SiteUrl
    Site collections to cover, instead of all of them.

.PARAMETER SiteFilter
    Only site collections whose title or address contains this text.

.PARAMETER OutputPath
    Folder for the CSV files. Default: the current folder.

.EXAMPLE
    ./Get-AccessReport.ps1 -AdminUrl https://contoso-admin.sharepoint.com -Subject 'Ann=ann@contoso.com'

.EXAMPLE
    ./Get-AccessReport.ps1 -AdminUrl https://contoso-admin.sharepoint.com `
        -Subject 'Auditors=c:0t.c|tenant|11111111-1111-1111-1111-111111111111', 'pat@fabrikam.com' `
        -SiteFilter finance -OutputPath C:\Reports
#>
[CmdletBinding()]
param(
    [string]$AdminUrl,
    [Parameter(Mandatory = $true)][string[]]$Subject,
    [string[]]$SiteUrl,
    [string]$SiteFilter = '',
    [string]$OutputPath = (Get-Location).Path
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ReportHeaders = @('Who', 'Site URL', 'Site title', 'Object', 'Name', 'URL', 'Access', 'How', 'Own permissions', 'Site collection')
$UnreadHeaders = @('Who', 'Site URL', 'Site title', 'Step', 'Status', 'Error')
$EveryoneClaim = 'c:0(.s|true'

# ---------------------------------------------------------------------------------------
# The report format: identical to the web part's (services/reportCsv.ts)

function ConvertTo-ReportCell([object]$Value) {
    $s = if ($null -eq $Value) { '' } else { [string]$Value }
    # A title such as =HYPERLINK(...) must not run as a formula in Excel.
    if ($s -cmatch '^[=+\-@\t\r]') { $s = "'" + $s }
    return '"' + $s.Replace('"', '""') + '"'
}

function ConvertTo-ReportCsv([string[]]$Headers, $Rows) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append([char]0xFEFF)
    $all = @(, $Headers) + @($Rows)
    foreach ($r in $all) {
        $cells = foreach ($c in $r) { ConvertTo-ReportCell $c }
        [void]$sb.Append(($cells -join ',')).Append("`r`n")
    }
    return $sb.ToString()
}

# ---------------------------------------------------------------------------------------
# Subjects

function ConvertTo-Login([string]$Value) {
    $v = $Value.Trim()
    if ($v.Contains('|')) {
        # A Microsoft 365 group is kept as its members claim; its owners claim is checked with it.
        if ($v.ToLowerInvariant().Contains('|federateddirectoryclaimprovider|')) { $v = $v -replace '_o$', '' }
        return $v
    }
    return 'i:0#.f|membership|' + $v.ToLowerInvariant()
}

$subjects = foreach ($x in $Subject) {
    if ($x -match '^([^=|]+)=(.+)$') { $label = $Matches[1].Trim(); $value = $Matches[2] }
    else { $label = $x.Trim(); $value = $x }
    $login = ConvertTo-Login $value
    $lower = $login.ToLowerInvariant()
    $isPerson = $lower.StartsWith('i:')
    $isM365 = $lower.Contains('|federateddirectoryclaimprovider|')
    # The logins this subject can appear under in a site: a Microsoft 365 group has members and owners.
    $claims = @(@{ Login = $login; Who = $label })
    if ($isM365) { $claims += @{ Login = $login + '_o'; Who = "$label (owners)" } }
    New-Object PSObject -Property @{ Label = $label; Login = $login; IsPerson = $isPerson; IsGuest = ($isPerson -and $lower.Contains('#ext#')); Claims = $claims }
}
$subjects = @($subjects)

# ---------------------------------------------------------------------------------------
# Connection and site collections

if ($AdminUrl) {
    Connect-SPOService -Url $AdminUrl.Trim().TrimEnd('/')
}

function Invoke-WithRetry([scriptblock]$Action) {
    for ($attempt = 0; ; $attempt++) {
        try { return & $Action }
        catch {
            if ($attempt -lt 4 -and [string]$_.Exception.Message -match '429|503|throttl|Too many requests|temporarily') {
                Start-Sleep -Seconds ([math]::Min(120, 5 * [math]::Pow(2, $attempt)))
                continue
            }
            throw
        }
    }
}

function Get-StatusOf($ErrorRecord) {
    $m = [regex]::Match([string]$ErrorRecord.Exception.Message, '\b([45]\d\d)\b')
    if ($m.Success) { return $m.Groups[1].Value }
    if ([string]$ErrorRecord.Exception.Message -match 'denied|unauthori[sz]ed') { return '403' }
    if ([string]$ErrorRecord.Exception.Message -match 'be found|not found|does not exist') { return '404' }
    return ''
}

function Get-MessageOf($ErrorRecord) {
    $m = ([string]$ErrorRecord.Exception.Message -replace '[\r\n]+', ' ').Trim()
    if ($m.Length -gt 300) { $m = $m.Substring(0, 300) }
    return $m
}

try {
    $sites = if ($SiteUrl) {
        foreach ($u in $SiteUrl) { Invoke-WithRetry { Get-SPOSite -Identity $u.Trim().TrimEnd('/') } }
    }
    else {
        Write-Host 'Listing site collections...'
        Invoke-WithRetry { Get-SPOSite -Limit All }
    }
}
catch {
    if ($_.Exception.Message -match 'connect|Connect-SPOService|No connection') {
        throw 'Not connected: pass -AdminUrl https://<tenant>-admin.sharepoint.com, or run Connect-SPOService first.'
    }
    throw
}

$filter = $SiteFilter.Trim().ToLowerInvariant()
$sites = @($sites | Where-Object {
        $url = ([string]$_.Url).TrimEnd('/')
        $template = [string]$_.Template
        $url -and
        $url -notmatch '^https://[^/]+/personal/' -and
        $template -notmatch '^SPSPERS' -and
        $template -notmatch '^REDIRECTSITE' -and
        ("$($_.Title) $url").ToLowerInvariant().Contains($filter)
    })

# ---------------------------------------------------------------------------------------
# Scan

$rows = New-Object System.Collections.ArrayList
$unread = New-Object System.Collections.ArrayList

function Add-Row($Who, $Site, $Level, $How) {
    $url = ([string]$Site.Url).TrimEnd('/')
    $title = if ($Site.Title) { [string]$Site.Title } else { $url }
    [void]$rows.Add(@($Who, $url, $title, 'Site', $title, $url, $Level, $How, '', $url))
}

function Add-Unread($Who, $Site, $Step, $ErrorRecord) {
    $url = ([string]$Site.Url).TrimEnd('/')
    $title = if ($Site.Title) { [string]$Site.Title } else { $url }
    [void]$unread.Add(@($Who, $url, $title, $Step, (Get-StatusOf $ErrorRecord), (Get-MessageOf $ErrorRecord)))
}

function Test-Contains([string[]]$Logins, [string]$Login) {
    foreach ($l in $Logins) { if ([string]::Equals([string]$l, $Login, [StringComparison]::OrdinalIgnoreCase)) { return $true } }
    return $false
}

Write-Host "Checking $($sites.Count) site collections for $($subjects.Count) people and groups (site collection level)."
$n = 0
foreach ($site in $sites) {
    $n++
    Write-Progress -Activity 'Access report' -Status "$n of $($sites.Count): $($site.Url)" -PercentComplete (100 * ($n - 1) / [math]::Max(1, $sites.Count))
    $url = ([string]$site.Url).TrimEnd('/')

    # The site's SharePoint groups, their members and permission levels: one call per site.
    try {
        $groups = @(Invoke-WithRetry { Get-SPOSiteGroup -Site $url -Limit 1000 })
    }
    catch {
        Add-Unread '(all)' $site 'SharePoint groups' $_
        continue
    }
    $groups = @($groups | ForEach-Object {
            $levels = @($_.Roles | Where-Object { $_ -and $_ -notmatch '^limited access' })
            New-Object PSObject -Property @{ Title = [string]$_.Title; Users = @($_.Users); Level = ($levels -join ', ') }
        } | Where-Object { $_.Level })

    foreach ($s in $subjects) {
        foreach ($c in $s.Claims) {
            # Site collection administrator?
            try {
                $user = Invoke-WithRetry { Get-SPOUser -Site $url -LoginName $c.Login }
                if ($user -and $user.IsSiteAdmin) { Add-Row $c.Who $site 'Full control' 'Site collection administrator' }
            }
            catch {
                # Not in the site's user list: not an administrator. Anything else could not be read.
                if ([string]$_.Exception.Message -notmatch 'be found|not found|does not exist') {
                    Add-Unread $c.Who $site 'Site collection administrators' $_
                }
            }
            # SharePoint groups it is a member of.
            foreach ($g in $groups) {
                if (Test-Contains $g.Users $c.Login) {
                    Add-Row $c.Who $site $g.Level "SharePoint group: $($g.Title)"
                }
            }
        }
        if ($s.IsPerson) {
            # Groups that let in everyone, or everyone in the organisation (not guests).
            foreach ($g in $groups) {
                if (Test-Contains $g.Users $s.Login) { continue }
                $everyone = Test-Contains $g.Users $EveryoneClaim
                $internal = -not $s.IsGuest -and @($g.Users | Where-Object { ([string]$_).ToLowerInvariant().StartsWith('c:0-.f|rolemanager|spo-grid-all-users/') }).Count -gt 0
                if ($everyone) { Add-Row $s.Label $site $g.Level "SharePoint group: $($g.Title) (via Everyone)" }
                elseif ($internal) { Add-Row $s.Label $site $g.Level "SharePoint group: $($g.Title) (via Everyone except external users)" }
            }
        }
    }
}
Write-Progress -Activity 'Access report' -Completed

# ---------------------------------------------------------------------------------------
# Files

$stamp = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd-HHmm', [Globalization.CultureInfo]::InvariantCulture)
$utf8 = New-Object System.Text.UTF8Encoding($false)
if (-not (Test-Path $OutputPath)) { [void](New-Item -ItemType Directory -Path $OutputPath) }
$outDir = (Resolve-Path $OutputPath).Path
$reportPath = Join-Path $outDir "access-report-$stamp.csv"
[System.IO.File]::WriteAllText($reportPath, (ConvertTo-ReportCsv $ReportHeaders $rows), $utf8)
$written = @($reportPath)
if ($unread.Count) {
    $unreadPath = Join-Path $outDir "access-report-$stamp-unread.csv"
    [System.IO.File]::WriteAllText($unreadPath, (ConvertTo-ReportCsv $UnreadHeaders $unread), $utf8)
    $written += $unreadPath
}
Write-Host "$($rows.Count) rows, $($unread.Count) site collections or steps not read."
$written | ForEach-Object { Get-Item $_ }
