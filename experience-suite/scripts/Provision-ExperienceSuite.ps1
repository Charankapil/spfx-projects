<#
.SYNOPSIS
  Creates the SharePoint lists and fields required by SharePoint Experience Suite.
.DESCRIPTION
  Run as a site owner/site collection administrator after installing PnP.PowerShell.
  This script does not grant permissions or create tenant-wide app registrations.
.PARAMETER ClientId
  Client ID of the Entra ID app used by PnP.PowerShell for interactive sign-in. Since September 2024
  PnP.PowerShell no longer ships a shared multi-tenant app, so -Interactive needs your own app
  (create one once with Register-PnPEntraIDAppForInteractiveLogin) unless the ENTRAID_APP_ID /
  ENTRAID_CLIENT_ID environment variable is set. This app is only used by the provisioning script;
  the web part itself still needs no app registration.
.EXAMPLE
  .\Provision-ExperienceSuite.ps1 -SiteUrl https://contoso.sharepoint.com/sites/EmployeeHub -ClientId 00000000-0000-0000-0000-000000000000
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [ValidatePattern('^https://')]
  [string]$SiteUrl,

  [Parameter(Mandatory = $false)]
  [string]$ClientId
)

$ErrorActionPreference = 'Stop'
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
  throw 'PnP.PowerShell is required. Install it with: Install-Module PnP.PowerShell -Scope CurrentUser'
}

if ($ClientId) {
  Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
} elseif ($env:ENTRAID_APP_ID -or $env:ENTRAID_CLIENT_ID) {
  Connect-PnPOnline -Url $SiteUrl -Interactive
} else {
  throw 'Provide -ClientId (or set ENTRAID_APP_ID). PnP.PowerShell requires your own Entra ID app for interactive sign-in; see Register-PnPEntraIDAppForInteractiveLogin.'
}

function Ensure-List([string]$Title) {
  $existing = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
  if (-not $existing) {
    Write-Host "Creating list: $Title" -ForegroundColor Cyan
    New-PnPList -Title $Title -Template GenericList -OnQuickLaunch:$false | Out-Null
  } else {
    Write-Host "List exists: $Title" -ForegroundColor DarkGray
  }
}

function Ensure-Field([string]$List, [string]$InternalName, [string]$DisplayName, [string]$Type, [string[]]$Choices = @()) {
  $existing = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue
  if ($existing) {
    if ($existing.InternalName -ne $InternalName) {
      Write-Warning "$List has a column titled '$InternalName' whose internal name is '$($existing.InternalName)'. The web part needs the internal name '$InternalName'."
    }
    return
  }
  Write-Host "  Adding field: $InternalName" -ForegroundColor Cyan
  if ($Type -eq 'Choice') {
    $choiceXml = ($Choices | ForEach-Object { "<CHOICE>$([System.Security.SecurityElement]::Escape($_))</CHOICE>" }) -join ''
    $default = [System.Security.SecurityElement]::Escape($Choices[0])
    $xml = "<Field Type='Choice' DisplayName='$DisplayName' Name='$InternalName' StaticName='$InternalName' Format='Dropdown'><Default>$default</Default><CHOICES>$choiceXml</CHOICES></Field>"
  } elseif ($Type -eq 'Note') {
    $xml = "<Field Type='Note' DisplayName='$DisplayName' Name='$InternalName' StaticName='$InternalName' NumLines='6' RichText='FALSE' />"
  } elseif ($Type -eq 'DateTime') {
    $xml = "<Field Type='DateTime' DisplayName='$DisplayName' Name='$InternalName' StaticName='$InternalName' Format='DateTime' />"
  } else {
    $xml = "<Field Type='$Type' DisplayName='$DisplayName' Name='$InternalName' StaticName='$InternalName' />"
  }
  Add-PnPFieldFromXml -List $List -FieldXml $xml | Out-Null
  # Fields created from XML are not added to any view; add them to the default view so list owners can see them.
  $view = Get-PnPView -List $List | Where-Object { $_.DefaultView } | Select-Object -First 1
  if ($view -and ($view.ViewFields -notcontains $InternalName)) {
    $fields = @($view.ViewFields) + $InternalName
    Set-PnPView -List $List -Identity $view.Id -Fields $fields | Out-Null
  }
}

$definitions = @(
  @{ Name='ES Requests'; Fields=@(
    @('Description','Description','Note'), @('Category','Category','Text'),
    @('Status','Status','Choice',@('Submitted','Assigned','In Progress','Completed','Rejected')),
    @('RequesterName','Requester Name','Text'), @('RequesterEmail','Requester Email','Text'),
    @('Owner','Owner','Text'), @('DueDate','Due Date','DateTime')
  )},
  @{ Name='ES Approvals'; Fields=@(
    @('Description','Description','Note'), @('Category','Category','Text'),
    @('Status','Status','Choice',@('Pending','Approved','Rejected')),
    @('Owner','Owner','Text'), @('ApproverEmail','Approver Email','Text'),
    @('DueDate','Due Date','DateTime'), @('DecisionBy','Decision By','Text'), @('DecisionDate','Decision Date','DateTime')
  )},
  @{ Name='ES Announcements'; Fields=@(
    @('Description','Description','Note'), @('Category','Category','Text'), @('Owner','Owner','Text')
  )},
  @{ Name='ES Tasks'; Fields=@(
    @('Description','Description','Note'), @('Category','Category','Text'),
    @('Status','Status','Choice',@('Not Started','In Progress','Completed')),
    @('Owner','Owner','Text'), @('DueDate','Due Date','DateTime')
  )}
)

foreach ($definition in $definitions) {
  Ensure-List -Title $definition.Name
  foreach ($field in $definition.Fields) {
    $choices = @()
    if ($field.Count -ge 4) { $choices = $field[3] }
    Ensure-Field -List $definition.Name -InternalName $field[0] -DisplayName $field[1] -Type $field[2] -Choices $choices
  }
}

Write-Host ''
Write-Host 'SharePoint Experience Suite lists and fields are ready.' -ForegroundColor Green
Write-Host 'Review list permissions and configure the ApproverEmail/Owner values before enabling approval actions.' -ForegroundColor Yellow
