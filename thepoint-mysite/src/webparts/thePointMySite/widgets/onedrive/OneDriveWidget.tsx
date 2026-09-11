import * as React from 'react';
import { Breadcrumb, IconButton, Pivot, PivotItem, ProgressIndicator, SearchBox, type IBreadcrumbItem } from '@fluentui/react';
import type { IWidgetDefinition, IWidgetProps } from '../../models';
import { useAsync } from '../../hooks/useAsync';
import { LoadingState, ErrorState, EmptyState, FooterLink } from '../../components/shared/States';
import { ItemList, ItemRow } from '../../components/shared/ItemRow';
import shared from '../../components/shared/Shared.module.scss';
import type { IGraphDrive, IGraphDriveItem } from '../../services/GraphService';
import { fileIcon, fileIconColor, formatBytes, formatRelative } from '../../utils/format';

interface IFolderCrumb {
  id: string | undefined;
  name: string;
}

const OneDriveWidget: React.FC<IWidgetProps> = ({ graph, settings, density }) => {
  const max: number = Number(settings.maxItems) || 15;
  const [view, setView] = React.useState<string>(String(settings.defaultView || 'recent'));
  const [path, setPath] = React.useState<IFolderCrumb[]>([{ id: undefined, name: 'My files' }]);
  const [query, setQuery] = React.useState<string>('');
  const [debounced, setDebounced] = React.useState<string>('');

  React.useEffect(() => {
    const t: number = window.setTimeout(() => setDebounced(query.trim()), 350);
    return () => window.clearTimeout(t);
  }, [query]);

  const drive = useAsync<IGraphDrive>(() => graph.getMyDrive(), [graph]);
  const current: IFolderCrumb = path[path.length - 1];

  const items = useAsync<IGraphDriveItem[]>(async () => {
    if (debounced) {
      return graph.searchMyDrive(debounced, max);
    }
    if (view === 'browse') {
      return graph.getDriveChildren(current.id, 100);
    }
    if (view === 'shared') {
      return graph.getSharedWithMe(max);
    }
    return graph.getRecentFiles(max);
  }, [graph, view, current.id, max, debounced]);

  const openFolder = (item: IGraphDriveItem): void => setPath(prev => [...prev, { id: item.id, name: item.name }]);

  const crumbs: IBreadcrumbItem[] = path.map((p, i) => ({
    key: p.id || 'root',
    text: p.name,
    onClick: i < path.length - 1 ? (): void => setPath(path.slice(0, i + 1)) : undefined
  }));

  const body: React.ReactNode = (() => {
    if (items.loading && !items.data) {
      return <LoadingState rows={5} />;
    }
    if (items.error) {
      return <ErrorState message={items.error} onRetry={items.reload} />;
    }
    if (!items.data || items.data.length === 0) {
      return <EmptyState icon="OneDriveLogo" title={debounced ? 'No matching files' : 'Nothing here yet'} />;
    }
    return (
      <ItemList compact={density === 'compact'}>
        {items.data.map(item => {
          const remote = item.remoteItem;
          const name: string = item.name || remote?.name || 'Untitled';
          const isFolder: boolean = !!item.folder;
          const icon: string = isFolder ? 'FabricFolder' : fileIcon(name);
          const modified: string | undefined = item.lastModifiedDateTime || remote?.lastModifiedDateTime;
          const sharedBy: string | undefined = item.shared?.sharedBy?.user?.displayName || remote?.shared?.sharedBy?.user?.displayName;
          const meta: string = [
            isFolder ? `${item.folder?.childCount ?? 0} items` : formatBytes(item.size),
            sharedBy ? `from ${sharedBy}` : item.lastModifiedBy?.user?.displayName,
            modified ? formatRelative(modified) : undefined
          ].filter(Boolean).join(' · ');
          const href: string | undefined = isFolder ? undefined : item.webUrl || remote?.webUrl;
          return (
            <ItemRow
              key={item.id}
              title={name}
              href={href}
              onClick={isFolder ? (): void => openFolder(item) : undefined}
              icon={icon}
              iconColor={fileIconColor(icon)}
              meta={meta}
              trailing={isFolder ? <IconButton iconProps={{ iconName: 'ChevronRight' }} styles={{ root: { height: 22, width: 22 }, icon: { fontSize: 10 } }} /> : undefined}
            />
          );
        })}
      </ItemList>
    );
  })();

  const quota = drive.data?.quota;
  const usedPct: number = quota && quota.total ? (quota.used || 0) / quota.total : 0;

  return (
    <div>
      {settings.showSearch !== false && (
        <div className={shared.toolbar}>
          <SearchBox placeholder="Search my OneDrive" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { flex: 1, height: 30 } }} />
        </div>
      )}
      {!debounced && (
        <Pivot selectedKey={view} onLinkClick={i => setView(i?.props.itemKey || 'recent')} headersOnly styles={{ link: { height: 30, fontSize: 12 }, linkIsSelected: { height: 30, fontSize: 12 } }}>
          <PivotItem headerText="Recent" itemKey="recent" />
          <PivotItem headerText="My files" itemKey="browse" />
          <PivotItem headerText="Shared" itemKey="shared" />
        </Pivot>
      )}
      {!debounced && view === 'browse' && (
        <Breadcrumb items={crumbs} maxDisplayedItems={3} styles={{ root: { margin: '4px 0' }, itemLink: { fontSize: 12, height: 26, lineHeight: 26 }, item: { fontSize: 12, height: 26, lineHeight: 26 } }} />
      )}
      <div style={{ paddingTop: 4 }}>{body}</div>
      {settings.showQuota !== false && quota && quota.total && (
        <div style={{ paddingTop: 8 }}>
          <ProgressIndicator
            percentComplete={usedPct}
            description={`${formatBytes(quota.used)} of ${formatBytes(quota.total)} used${quota.state && quota.state !== 'normal' ? ` · ${quota.state}` : ''}`}
            styles={{ itemDescription: { fontSize: 11.5 } }}
          />
        </div>
      )}
      {drive.data?.webUrl && <FooterLink href={drive.data.webUrl} text="Open OneDrive" />}
    </div>
  );
};

export const oneDriveDefinition: IWidgetDefinition = {
  type: 'onedrive',
  title: 'My OneDrive',
  description: 'Recent files, folder browsing, files shared with you, search and storage usage for your OneDrive.',
  icon: 'OneDriveLogo',
  category: 'content',
  defaultSize: 1,
  allowMultiple: false,
  permissions: ['Files.Read.All'],
  settingsSchema: [
    { key: 'maxItems', label: 'Number of files', type: 'number', defaultValue: 15, min: 5, max: 50 },
    { key: 'defaultView', label: 'Default view', type: 'choice', defaultValue: 'recent', choices: [{ key: 'recent', text: 'Recent' }, { key: 'browse', text: 'My files' }, { key: 'shared', text: 'Shared' }] },
    { key: 'showSearch', label: 'Show search box', type: 'boolean', defaultValue: true },
    { key: 'showQuota', label: 'Show storage usage', type: 'boolean', defaultValue: true }
  ],
  component: OneDriveWidget
};

export default OneDriveWidget;
