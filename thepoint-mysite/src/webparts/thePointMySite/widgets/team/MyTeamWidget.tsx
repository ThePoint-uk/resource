import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react';
import type { IWidgetDefinition, IWidgetProps } from '../../models';
import { useAsync } from '../../hooks/useAsync';
import { LoadingState, ErrorState, EmptyState, SectionHeader } from '../../components/shared/States';
import { ItemList } from '../../components/shared/ItemRow';
import { PersonRow } from './PersonRow';
import type { IGraphPresence, IGraphUser } from '../../services/GraphService';

interface ITeamData {
  manager: IGraphUser | undefined;
  peers: IGraphUser[];
  reports: IGraphUser[];
  meId: string;
}

const MyTeamWidget: React.FC<IWidgetProps> = ({ graph, settings, context, density }) => {
  const showPresence: boolean = settings.showPresence !== false;
  const maxPeople: number = Number(settings.maxPeople) || 12;
  const defaultView: string = String(settings.defaultView || 'all');
  const [view, setView] = React.useState<string>(defaultView);

  const team = useAsync<ITeamData>(async () => {
    const me: { id: string } = await graph.get<{ id: string }>('/me?$select=id');
    const [manager, reports] = await Promise.all([graph.getMyManager(), graph.getDirectReports('me')]);
    let peers: IGraphUser[] = [];
    if (manager) {
      peers = (await graph.getDirectReports(manager.id)).filter(p => p.id !== me.id);
    }
    return { manager, peers, reports, meId: me.id };
  }, [graph]);

  const ids: string[] = React.useMemo(() => {
    if (!team.data) {
      return [];
    }
    return [team.data.manager, ...team.data.peers, ...team.data.reports].filter((u): u is IGraphUser => !!u).map(u => u.id);
  }, [team.data]);

  const presence = useAsync<Map<string, string>>(async () => {
    const map: Map<string, string> = new Map();
    if (!showPresence || ids.length === 0) {
      return map;
    }
    try {
      const result: IGraphPresence[] = await graph.getPresence(ids);
      result.forEach(p => map.set(p.id, p.availability));
    } catch {
      /* presence is optional */
    }
    return map;
  }, [graph, ids.join(','), showPresence]);

  if (team.loading && !team.data) {
    return <LoadingState rows={4} />;
  }
  if (team.error || !team.data) {
    return <ErrorState message={team.error || 'Team unavailable'} onRetry={team.reload} />;
  }
  const { manager, peers, reports } = team.data;
  if (!manager && peers.length === 0 && reports.length === 0) {
    return <EmptyState icon="Group" title="No team information" message="Your manager and direct reports have not been set in the directory yet." />;
  }

  const compact: boolean = density === 'compact';
  const p = (u: IGraphUser): string | undefined => presence.data?.get(u.id);
  const showTabs: boolean = settings.showTabs !== false;
  const sections: { key: string; title: string; people: IGraphUser[]; badge?: string }[] = [
    { key: 'manager', title: 'Manager', people: manager ? [manager] : [], badge: 'Manager' },
    { key: 'peers', title: 'Colleagues', people: peers },
    { key: 'reports', title: 'Direct reports', people: reports }
  ].filter(s => s.people.length > 0);

  const visible = showTabs && view !== 'all' ? sections.filter(s => s.key === view) : sections;

  return (
    <div>
      {showTabs && sections.length > 1 && (
        <Pivot selectedKey={view} onLinkClick={i => setView(i?.props.itemKey || 'all')} headersOnly styles={{ link: { height: 30, fontSize: 12 }, linkIsSelected: { height: 30, fontSize: 12 } }}>
          <PivotItem headerText="All" itemKey="all" />
          {sections.map(s => <PivotItem headerText={`${s.title} (${s.people.length})`} itemKey={s.key} key={s.key} />)}
        </Pivot>
      )}
      {visible.map(section => (
        <div key={section.key}>
          {(view === 'all' || !showTabs) && <SectionHeader text={`${section.title} · ${section.people.length}`} />}
          <ItemList compact={compact}>
            {section.people.slice(0, view === 'all' ? maxPeople : 200).map(u => (
              <PersonRow key={u.id} graph={graph} user={u} presence={p(u)} badge={section.badge} compact={compact} />
            ))}
          </ItemList>
          {view === 'all' && section.people.length > maxPeople && (
            <div style={{ fontSize: 12, padding: '4px 4px 0', color: 'var(--tpms-text-muted)' }}>
              +{section.people.length - maxPeople} more · <a href="#" onClick={e => { e.preventDefault(); setView(section.key); }} style={{ color: 'var(--tpms-accent)' }}>show all</a>
            </div>
          )}
        </div>
      ))}
    </div>
  );
};

export const myTeamDefinition: IWidgetDefinition = {
  type: 'team',
  title: 'My team',
  description: 'Your manager, colleagues who share your manager and your direct reports, with presence and quick actions.',
  icon: 'Group',
  category: 'people',
  defaultSize: 1,
  allowMultiple: false,
  permissions: ['User.Read.All', 'Presence.Read.All'],
  settingsSchema: [
    { key: 'showPresence', label: 'Show Teams presence', type: 'boolean', defaultValue: true },
    { key: 'showTabs', label: 'Show tabs', type: 'boolean', defaultValue: true },
    { key: 'maxPeople', label: 'People per section (All view)', type: 'number', defaultValue: 8, min: 3, max: 50 },
    {
      key: 'defaultView',
      label: 'Default view',
      type: 'choice',
      defaultValue: 'all',
      choices: [
        { key: 'all', text: 'All' },
        { key: 'manager', text: 'Manager' },
        { key: 'peers', text: 'Colleagues' },
        { key: 'reports', text: 'Direct reports' }
      ]
    }
  ],
  component: MyTeamWidget
};

export default MyTeamWidget;
