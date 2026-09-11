import * as React from 'react';
import { Pivot, PivotItem, SearchBox } from '@fluentui/react';
import type { IWidgetDefinition, IWidgetProps } from '../../models';
import { useAsync } from '../../hooks/useAsync';
import { LoadingState, ErrorState, EmptyState } from '../../components/shared/States';
import { ItemList, ItemRow } from '../../components/shared/ItemRow';
import shared from '../../components/shared/Shared.module.scss';
import type { IGraphSite } from '../../services/GraphService';
import type { ISPSite } from '../../services/SharePointService';

interface ISiteItem {
  id: string;
  title: string;
  url: string;
  description?: string;
  logoUrl?: string;
}

const SitesWidget: React.FC<IWidgetProps> = ({ graph, sp, settings, density }) => {
  const max: number = Number(settings.maxItems) || 10;
  const [view, setView] = React.useState<string>(String(settings.defaultView || 'followed'));
  const [query, setQuery] = React.useState<string>('');
  const [debounced, setDebounced] = React.useState<string>('');

  React.useEffect(() => {
    const t: number = window.setTimeout(() => setDebounced(query.trim()), 350);
    return () => window.clearTimeout(t);
  }, [query]);

  const sites = useAsync<ISiteItem[]>(async () => {
    if (debounced) {
      const results: ISPSite[] = await sp.getSites(max, debounced);
      return results.map(s => ({ id: s.id, title: s.title, url: s.url, description: s.description, logoUrl: s.logoUrl }));
    }
    if (view === 'followed') {
      const followed: IGraphSite[] = await graph.getFollowedSites(max);
      return followed.map(s => ({ id: s.id, title: s.displayName || s.name || s.webUrl, url: s.webUrl, description: s.description }));
    }
    const all: ISPSite[] = await sp.getSites(max);
    return all.map(s => ({ id: s.id, title: s.title, url: s.url, description: s.description, logoUrl: s.logoUrl }));
  }, [graph, sp, view, max, debounced]);

  const layout: string = String(settings.layout || 'list');
  const newTab: boolean = settings.openInNewTab !== false;

  const body: React.ReactNode = (() => {
    if (sites.loading && !sites.data) {
      return <LoadingState rows={5} />;
    }
    if (sites.error) {
      return <ErrorState message={sites.error} onRetry={sites.reload} />;
    }
    if (!sites.data || sites.data.length === 0) {
      return (
        <EmptyState icon="SharepointLogo" title={debounced ? 'No matching sites' : view === 'followed' ? 'You are not following any sites' : 'No sites found'} message={debounced ? undefined : view === 'followed' ? 'Use the star on any SharePoint site to follow it.' : undefined} />
      );
    }
    if (layout === 'tiles') {
      return (
        <div className={shared.grid}>
          {sites.data.map(s => (
            <a key={s.id} className={shared.tile} href={s.url} target={newTab ? '_blank' : undefined} rel="noreferrer" title={s.description || s.title}>
              <span className={shared.tileIcon} style={{ background: 'var(--tpms-accent-soft)', color: 'var(--tpms-accent)', fontWeight: 700, fontSize: 14 }}>
                {s.logoUrl ? <img src={s.logoUrl} alt="" loading="lazy" /> : (s.title || '?').slice(0, 2).toUpperCase()}
              </span>
              <span className={shared.tileTitle}>{s.title}</span>
            </a>
          ))}
        </div>
      );
    }
    return (
      <ItemList compact={density === 'compact'}>
        {sites.data.map(s => (
          <ItemRow
            key={s.id}
            title={s.title}
            href={s.url}
            newTab={newTab}
            imageUrl={s.logoUrl}
            icon="SharepointLogo"
            iconColor="#038387"
            meta={s.description || s.url.replace(/^https?:\/\//, '')}
          />
        ))}
      </ItemList>
    );
  })();

  return (
    <div>
      {settings.showSearch !== false && (
        <div className={shared.toolbar}>
          <SearchBox placeholder="Search sites" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { flex: 1, height: 30 } }} />
        </div>
      )}
      {!debounced && settings.showTabs !== false && (
        <Pivot selectedKey={view} onLinkClick={i => setView(i?.props.itemKey || 'followed')} headersOnly styles={{ link: { height: 30, fontSize: 12 }, linkIsSelected: { height: 30, fontSize: 12 } }}>
          <PivotItem headerText="Following" itemKey="followed" />
          <PivotItem headerText="Recently active" itemKey="all" />
        </Pivot>
      )}
      <div style={{ paddingTop: 4 }}>{body}</div>
    </div>
  );
};

export const sitesDefinition: IWidgetDefinition = {
  type: 'sites',
  title: 'Sites',
  description: 'SharePoint sites you follow, recently active sites you can access, and site search.',
  icon: 'SharepointLogo',
  category: 'content',
  defaultSize: 1,
  allowMultiple: true,
  permissions: ['Sites.Read.All'],
  settingsSchema: [
    { key: 'maxItems', label: 'Number of sites', type: 'number', defaultValue: 10, min: 3, max: 40 },
    { key: 'layout', label: 'Layout', type: 'choice', defaultValue: 'list', choices: [{ key: 'list', text: 'List' }, { key: 'tiles', text: 'Tiles' }] },
    { key: 'defaultView', label: 'Default view', type: 'choice', defaultValue: 'followed', choices: [{ key: 'followed', text: 'Following' }, { key: 'all', text: 'Recently active' }] },
    { key: 'showTabs', label: 'Show tabs', type: 'boolean', defaultValue: true },
    { key: 'showSearch', label: 'Show search box', type: 'boolean', defaultValue: true },
    { key: 'openInNewTab', label: 'Open sites in a new tab', type: 'boolean', defaultValue: true }
  ],
  component: SitesWidget
};

export default SitesWidget;
