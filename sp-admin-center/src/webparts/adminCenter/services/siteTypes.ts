/** Friendly name for a site template code (STS#3, GROUP#0, SITEPAGEPUBLISHING#0, ...). */
export function siteType(template: string, groupConnected: boolean): string {
  const t = (template || '').toUpperCase();
  if (t.indexOf('SITEPAGEPUBLISHING') === 0) {
    return 'Communication site';
  }
  if (t.indexOf('GROUP') === 0 || (t.indexOf('STS#3') === 0 && groupConnected)) {
    return 'Team site (M365 group)';
  }
  if (t.indexOf('STS#3') === 0) {
    return 'Team site (no group)';
  }
  if (t.indexOf('STS') === 0) {
    return 'Classic team site';
  }
  if (t.indexOf('TEAMCHANNEL') === 0) {
    return 'Teams channel site';
  }
  if (t.indexOf('SPSPERS') === 0) {
    return 'OneDrive';
  }
  if (t.indexOf('APPCATALOG') === 0) {
    return 'App catalog';
  }
  if (t.indexOf('SRCHCEN') === 0 || t.indexOf('SRCHCENTERLITE') === 0) {
    return 'Search center';
  }
  if (t.indexOf('BLANKINTERNET') === 0 || t.indexOf('CMSPUBLISHING') === 0 || t.indexOf('ENTERWIKI') === 0) {
    return 'Classic publishing site';
  }
  if (t.indexOf('PROJECTSITE') === 0 || t.indexOf('PWA') === 0) {
    return 'Project site';
  }
  if (t.indexOf('REDIRECTSITE') === 0) {
    return 'Redirect site';
  }
  return template ? template : 'Unknown';
}

/** Is a "Teams connected" style cell truthy? */
export function isTruthy(raw: string | undefined): boolean {
  const v = (raw || '').trim().toLowerCase();
  return v === 'true' || v === 'yes' || v === '1' || v === 'y';
}
