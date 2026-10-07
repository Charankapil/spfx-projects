# SharePoint site ownership report

`Get-SiteOwnership.ps1` lists the real owners and members of every site collection and subsite, including M365 group owners and members. It needs no app registration.

## What counts as an owner
- Members of the web's associated **Owners** group
- Any principal with **Full Control** on the web (SharePoint group or direct)
- **M365 group owners** for group-connected sites

Site collection admins and the "primary admin" field are **excluded**, as are system accounts and Company Administrator.

## Prerequisites
```powershell
Install-Module PnP.PowerShell, ExchangeOnlineManagement -Scope CurrentUser
# only for -ResolveSecurityGroups:
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
```

## Usage
```powershell
# Try one site first
./Get-SiteOwnership.ps1 -TenantName contoso -SiteUrlFilter '*/sites/hr' -GrantTemporaryAccess

# Whole tenant, with visitors and Entra security group expansion
./Get-SiteOwnership.ps1 -TenantName contoso -GrantTemporaryAccess -IncludeVisitors -ResolveSecurityGroups
```

## Outputs
- `SiteOwnership_<ts>.csv`: one row per site, web, role and person
- `SitesWithoutOwners_<ts>.csv`: sites where no owner could be resolved (orphaned)
- `SiteOwnershipErrors_<ts>.csv`: sites or subsites that could not be read

## Notes
- **Sign-in:** the default is the Microsoft first-party SharePoint Online Management Shell app with device-code login. If your tenant blocks it, pass `-ClientId` for an app you are allowed to use.
- **-GrantTemporaryAccess** adds you as site admin for the duration of the read and removes you afterwards. These changes are audited. Without it, sites you cannot open show up in the errors file.
- **Unexpanded groups:** Entra security groups are reported as `UnexpandedSecurityGroup` unless you use `-ResolveSecurityGroups`, which needs Graph consent.
- **Status:** not run against a live tenant. Test on one site first.
