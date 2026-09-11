import * as React from 'react';
import { SearchBox } from '@fluentui/react';
import type { IWidgetDefinition, IWidgetProps } from '../../models';
import { useAsync } from '../../hooks/useAsync';
import { LoadingState, ErrorState, EmptyState, FooterLink } from '../../components/shared/States';
import { ItemList, ItemRow } from '../../components/shared/ItemRow';
import shared from '../../components/shared/Shared.module.scss';
import type { IGraphTeam } from '../../services/GraphService';

const TeamsWidget: React.FC<IWidgetProps> = ({ graph, settings, density }) => {
  const teams = useAsync<IGraphTeam[]>(async () => (await graph.getJoinedTeams()).filter(t => !t.isArchived || settings.showArchived === true).sort((a, b) => a.displayName.localeCompare(b.displayName)), [graph, settings.showArchived]);
  const [query, setQuery] = React.useState<string>('');
  const q: string = query.trim().toLowerCase();
  const list: IGraphTeam[] = (teams.data || []).filter(t => !q || t.displayName.toLowerCase().indexOf(q) >= 0);
  const max: number = Number(settings.maxItems) || 12;

  return (
    <div>
      {settings.showSearch !== false && (teams.data?.length || 0) > 5 && (
        <div className={shared.toolbar}>
          <SearchBox placeholder="Find a team" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { flex: 1, height: 30 } }} />
        </div>
      )}
      {teams.loading && !teams.data && <LoadingState rows={4} />}
      {teams.error && <ErrorState message={teams.error} onRetry={teams.reload} />}
      {teams.data && list.length === 0 && <EmptyState icon="TeamsLogo" title="No teams" message={q ? 'No teams match your search.' : 'You are not a member of any Teams yet.'} />}
      {list.length > 0 && (
        <ItemList compact={density === 'compact'}>
          {list.slice(0, q ? 100 : max).map(t => (
            <ItemRow
              key={t.id}
              title={t.displayName}
              href={t.webUrl || `https://teams.microsoft.com/l/team/${encodeURIComponent(t.id)}/conversations?groupId=${t.id}&tenantId=${graph.tenantId}`}
              leading={
                <span className={shared.rowIcon} style={{ background: '#6264a7', color: '#fff', fontWeight: 700, fontSize: 12 }}>
                  {t.displayName.split(/\s+/).slice(0, 2).map(p => p[0]).join('').toUpperCase()}
                </span>
              }
              meta={t.description || (t.isArchived ? 'Archived' : undefined)}
            />
          ))}
        </ItemList>
      )}
      <FooterLink href="https://teams.microsoft.com/" text="Open Teams" />
    </div>
  );
};

export const teamsDefinition: IWidgetDefinition = {
  type: 'teams',
  title: 'My Teams',
  description: 'Microsoft Teams you belong to, with deep links straight into each team.',
  icon: 'TeamsLogo',
  category: 'people',
  defaultSize: 1,
  allowMultiple: false,
  permissions: ['Team.ReadBasic.All'],
  settingsSchema: [
    { key: 'maxItems', label: 'Teams to show', type: 'number', defaultValue: 12, min: 3, max: 50 },
    { key: 'showSearch', label: 'Show search box', type: 'boolean', defaultValue: true },
    { key: 'showArchived', label: 'Include archived teams', type: 'boolean', defaultValue: false }
  ],
  component: TeamsWidget
};

export default TeamsWidget;
