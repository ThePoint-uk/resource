import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react';
import type { IWidgetDefinition, IWidgetProps } from '../../models';
import { useAsync } from '../../hooks/useAsync';
import { LoadingState, ErrorState, EmptyState, FooterLink } from '../../components/shared/States';
import { ItemList, ItemRow } from '../../components/shared/ItemRow';
import type { IGraphInsight } from '../../services/GraphService';
import { fileIcon, fileIconColor, formatRelative } from '../../utils/format';

const RecentDocumentsWidget: React.FC<IWidgetProps> = ({ graph, settings, density }) => {
  const max: number = Number(settings.maxItems) || 10;
  const showThumbnails: boolean = settings.showThumbnails === true;
  const [view, setView] = React.useState<string>(String(settings.defaultView || 'used'));

  const docs = useAsync<IGraphInsight[]>(async () => {
    switch (view) {
      case 'trending':
        return graph.getTrendingInsights(max);
      case 'shared':
        return graph.getSharedInsights(max);
      default:
        return graph.getUsedInsights(max);
    }
  }, [graph, view, max]);

  const body: React.ReactNode = (() => {
    if (docs.loading && !docs.data) {
      return <LoadingState rows={5} />;
    }
    if (docs.error) {
      return <ErrorState message={docs.error} onRetry={docs.reload} />;
    }
    if (!docs.data || docs.data.length === 0) {
      return <EmptyState icon="DocumentSet" title="No documents yet" message="Documents you open, edit or receive will show up here." />;
    }
    return (
      <ItemList compact={density === 'compact'}>
        {docs.data.map(d => {
          const icon: string = fileIcon(d.resourceVisualization.type || d.resourceVisualization.title);
          const when: string | undefined = d.lastUsed?.lastAccessedDateTime || d.lastUsed?.lastModifiedDateTime;
          return (
            <ItemRow
              key={d.id}
              title={d.resourceVisualization.title}
              href={d.resourceReference.webUrl}
              icon={icon}
              iconColor={fileIconColor(icon)}
              imageUrl={showThumbnails ? d.resourceVisualization.previewImageUrl : undefined}
              meta={[d.resourceVisualization.containerDisplayName, when ? formatRelative(when) : undefined].filter(Boolean).join(' · ')}
            />
          );
        })}
      </ItemList>
    );
  })();

  return (
    <div>
      {settings.showTabs !== false && (
        <Pivot selectedKey={view} onLinkClick={i => setView(i?.props.itemKey || 'used')} headersOnly styles={{ link: { height: 30, fontSize: 12 }, linkIsSelected: { height: 30, fontSize: 12 } }}>
          <PivotItem headerText="Recent" itemKey="used" />
          <PivotItem headerText="Shared with me" itemKey="shared" />
          <PivotItem headerText="Trending" itemKey="trending" />
        </Pivot>
      )}
      <div style={{ paddingTop: 4 }}>{body}</div>
      <FooterLink href="https://www.microsoft365.com/?auth=2#/mycontent" text="See all my content" />
    </div>
  );
};

export const recentDocumentsDefinition: IWidgetDefinition = {
  type: 'recentDocs',
  title: 'Recent documents',
  description: 'Files you have recently used, files shared with you and documents trending around you.',
  icon: 'DocumentSet',
  category: 'content',
  defaultSize: 1,
  allowMultiple: false,
  permissions: ['Sites.Read.All'],
  settingsSchema: [
    { key: 'maxItems', label: 'Number of items', type: 'number', defaultValue: 10, min: 3, max: 30 },
    { key: 'showTabs', label: 'Show Recent / Shared / Trending tabs', type: 'boolean', defaultValue: true },
    { key: 'defaultView', label: 'Default view', type: 'choice', defaultValue: 'used', choices: [{ key: 'used', text: 'Recent' }, { key: 'shared', text: 'Shared with me' }, { key: 'trending', text: 'Trending' }] },
    { key: 'showThumbnails', label: 'Show thumbnails', type: 'boolean', defaultValue: false }
  ],
  component: RecentDocumentsWidget
};

export default RecentDocumentsWidget;
