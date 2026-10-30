import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import { IBannerLock, matchLock } from '../../shared/readOnlyLocks';

export interface IReadOnlyBannerProperties {
  /** Written by the SharePoint Admin Center web part when something on this site is locked. */
  locks?: IBannerLock[];
}

const BANNER_ID = 'spac-readonly-banner';

/**
 * Shows the read-only message at the top of every page, list or library view that
 * falls inside a locked site, library or folder. It reads its lock list from its own
 * properties (set by the web part), so it makes no requests of its own.
 */
export default class ReadOnlyBannerApplicationCustomizer extends BaseApplicationCustomizer<IReadOnlyBannerProperties> {
  private top: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(this, this.render);
    // Modern pages navigate without a full reload; re-check the location each time.
    this.context.application.navigatedEvent.add(this, this.render);
    this.render();
    return Promise.resolve();
  }

  private render(): void {
    if (!this.top) {
      this.top = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, { onDispose: () => (this.top = undefined) });
    }
    if (!this.top || !this.top.domElement) {
      return;
    }
    const host = this.top.domElement;
    const lock = matchLock((this.properties && this.properties.locks) || [], window.location.pathname, window.location.search);
    const existing = host.querySelector(`#${BANNER_ID}`);
    if (!lock) {
      if (existing) {
        existing.remove();
      }
      return;
    }
    const banner = (existing as HTMLDivElement) || document.createElement('div');
    banner.id = BANNER_ID;
    banner.setAttribute('role', 'status');
    banner.style.cssText =
      'display:flex;gap:12px;align-items:flex-start;padding:10px 20px;background:#fff4ce;border-bottom:1px solid #f2c94c;color:#3b2a00;font:14px/1.45 "Segoe UI",-apple-system,Roboto,sans-serif;';
    banner.innerHTML = '';

    const icon = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    icon.setAttribute('viewBox', '0 0 24 24');
    icon.setAttribute('width', '20');
    icon.setAttribute('height', '20');
    icon.setAttribute('aria-hidden', 'true');
    icon.style.flex = 'none';
    icon.style.marginTop = '1px';
    const p = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    p.setAttribute('fill', '#8a5a00');
    p.setAttribute('d', 'M12 1a5 5 0 0 0-5 5v3H6a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-9a2 2 0 0 0-2-2h-1V6a5 5 0 0 0-5-5Zm-3 8V6a3 3 0 1 1 6 0v3H9Zm3 4a2 2 0 0 1 1 3.73V19h-2v-2.27A2 2 0 0 1 12 13Z');
    icon.appendChild(p);

    const text = document.createElement('div');
    const strong = document.createElement('strong');
    strong.textContent = lock.scope === 'web' ? 'This site is read-only. ' : lock.scope === 'list' ? 'This library is read-only. ' : 'This folder is read-only. ';
    const msg = document.createElement('span');
    // textContent, never innerHTML: the message is plain text typed by an admin.
    msg.textContent = lock.message || '';
    msg.style.whiteSpace = 'pre-wrap';
    text.appendChild(strong);
    text.appendChild(msg);

    banner.appendChild(icon);
    banner.appendChild(text);
    if (!existing) {
      host.appendChild(banner);
    }
  }
}
