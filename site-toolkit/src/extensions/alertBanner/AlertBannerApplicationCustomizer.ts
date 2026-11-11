import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';

import { alertStatus, dismiss, IAlert, isDismissed, loadAlerts } from '../../common/alerts';
import { safeHref } from '../../common/format';
import { SpClient } from '../../common/SpClient';
import styles from '../../common/ui/Toolkit.module.scss';

const SEVERITY_ORDER: Record<string, number> = { Critical: 0, Warning: 1, Info: 2 };
const MAX_BANNERS = 3;

/**
 * Shows the active alerts from the "Site Alerts" list at the top of every
 * page. Built with plain DOM calls (no React) to keep it tiny, and with
 * textContent only, so list content can never run as markup.
 */
export default class AlertBannerApplicationCustomizer extends BaseApplicationCustomizer<Record<string, never>> {
  private sp: SpClient | undefined;
  private placeholder: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
    this.sp = SpClient.fromContext(this.context);
    this.context.placeholderProvider.changedEvent.add(this, this.renderBanner);
    // In-site navigation swaps the page without reloading; the cached alerts are simply drawn again.
    this.context.application.navigatedEvent.add(this, this.renderBanner);
    this.renderBanner();
    return Promise.resolve();
  }

  private renderBanner(): void {
    const sp = this.sp;
    if (!sp) {
      return;
    }
    loadAlerts(sp)
      .then((alerts) => this.draw(sp, alerts || []))
      .catch(() => {
        // A banner must never break or clutter the page: on any failure it stays hidden.
      });
  }

  private draw(sp: SpClient, alerts: IAlert[]): void {
    if (!this.placeholder || !this.placeholder.domElement || !this.placeholder.domElement.isConnected) {
      this.placeholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
    }
    const host = this.placeholder ? this.placeholder.domElement : undefined;
    if (!host) {
      return;
    }
    while (host.firstChild) {
      host.removeChild(host.firstChild);
    }

    const visible = alerts
      .filter((alert) => alertStatus(alert) === 'active' && !isDismissed(sp, alert))
      .sort((a, b) => (SEVERITY_ORDER[a.Severity || 'Info'] ?? 2) - (SEVERITY_ORDER[b.Severity || 'Info'] ?? 2) || b.Id - a.Id)
      .slice(0, MAX_BANNERS);

    for (const alert of visible) {
      host.appendChild(this.buildBanner(sp, alert));
    }
  }

  private buildBanner(sp: SpClient, alert: IAlert): HTMLElement {
    const severity = alert.Severity === 'Critical' ? 'Critical' : alert.Severity === 'Warning' ? 'Warning' : 'Info';
    const tone = severity === 'Critical' ? styles.bannerCritical : severity === 'Warning' ? styles.bannerWarning : styles.bannerInfo;

    const banner = document.createElement('div');
    banner.className = `${styles.banner} ${tone}`;
    // Critical alerts interrupt screen readers; the others wait their turn.
    banner.setAttribute('role', severity === 'Critical' ? 'alert' : 'status');

    const text = document.createElement('div');
    text.className = styles.bannerText;
    const headline = document.createElement('strong');
    headline.textContent = alert.Title || severity;
    text.appendChild(headline);
    if (alert.Message) {
      text.appendChild(document.createTextNode(` ${alert.Message}`));
    }
    const href = safeHref(alert.LinkUrl);
    if (href) {
      text.appendChild(document.createTextNode(' '));
      const link = document.createElement('a');
      link.href = href;
      link.textContent = 'More information';
      link.style.color = 'inherit';
      link.style.fontWeight = '600';
      text.appendChild(link);
    }
    banner.appendChild(text);

    const close = document.createElement('button');
    close.type = 'button';
    close.className = styles.bannerClose;
    close.setAttribute('aria-label', `Dismiss: ${alert.Title || severity}`);
    close.textContent = '✕';
    close.addEventListener('click', () => {
      dismiss(sp, alert);
      if (banner.parentNode) {
        banner.parentNode.removeChild(banner);
      }
    });
    banner.appendChild(close);
    return banner;
  }
}
