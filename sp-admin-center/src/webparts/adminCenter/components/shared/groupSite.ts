import { graphErrorText } from './graph';
import { IAdminContext } from './context';

/**
 * Finds the SharePoint site of a Microsoft 365 group (one search request) and makes it the
 * site every page works on.
 */
export async function manageGroupSite(ctx: IAdminContext, groupId: string): Promise<void> {
  try {
    const res = await ctx.search.run(ctx.homeWebUrl, { query: `contentclass:STS_Site GroupId:${groupId}`, select: ['Path', 'Title'], rowLimit: 1 });
    const url = res.rows[0] && res.rows[0].Path;
    if (!url) {
      ctx.notify('The site for this group was not found in search (it may not be indexed yet or you have no access to it).', 'info');
      return;
    }
    const [web, site] = await Promise.all([ctx.api.getWeb(url), ctx.api.getSite(url)]);
    ctx.setTarget({ webUrl: web.url, siteUrl: site.url, title: web.title });
    ctx.notify(`Now managing ${web.title}.`, 'success');
  } catch (e) {
    ctx.notify(graphErrorText(e), 'error');
  }
}
