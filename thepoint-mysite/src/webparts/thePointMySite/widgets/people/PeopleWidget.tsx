import * as React from 'react';
import type { IWidgetDefinition, IWidgetProps } from '../../models';
import { useAsync } from '../../hooks/useAsync';
import { LoadingState, ErrorState, EmptyState } from '../../components/shared/States';
import { ItemList } from '../../components/shared/ItemRow';
import { PersonRow } from '../team/PersonRow';
import type { IGraphPerson, IGraphPresence, IGraphUser } from '../../services/GraphService';

const PeopleWidget: React.FC<IWidgetProps> = ({ graph, settings, density }) => {
  const max: number = Number(settings.maxPeople) || 8;
  const showPresence: boolean = settings.showPresence !== false;

  const people = useAsync<IGraphUser[]>(async () => {
    const result: IGraphPerson[] = await graph.getPeople(max);
    return result.map(p => ({
      id: p.id,
      displayName: p.displayName,
      jobTitle: p.jobTitle,
      department: p.department,
      mail: p.scoredEmailAddresses?.[0]?.address,
      userPrincipalName: p.userPrincipalName || p.scoredEmailAddresses?.[0]?.address
    }));
  }, [graph, max]);

  const presence = useAsync<Map<string, string>>(async () => {
    const map: Map<string, string> = new Map();
    if (!showPresence || !people.data || people.data.length === 0) {
      return map;
    }
    try {
      const result: IGraphPresence[] = await graph.getPresence(people.data.map(p => p.id));
      result.forEach(p => map.set(p.id, p.availability));
    } catch {
      /* optional */
    }
    return map;
  }, [graph, people.data, showPresence]);

  if (people.loading && !people.data) {
    return <LoadingState rows={4} />;
  }
  if (people.error || !people.data) {
    return <ErrorState message={people.error || 'People unavailable'} onRetry={people.reload} />;
  }
  if (people.data.length === 0) {
    return <EmptyState icon="People" title="No suggestions yet" message="People you work with will appear here as you collaborate." />;
  }
  return (
    <ItemList compact={density === 'compact'}>
      {people.data.map(u => (
        <PersonRow key={u.id} graph={graph} user={u} presence={presence.data?.get(u.id)} compact={density === 'compact'} />
      ))}
    </ItemList>
  );
};

export const peopleDefinition: IWidgetDefinition = {
  type: 'people',
  title: 'People I work with',
  description: 'The colleagues you collaborate with most, ranked by Microsoft 365 relevance.',
  icon: 'People',
  category: 'people',
  defaultSize: 1,
  allowMultiple: false,
  permissions: ['People.Read', 'Presence.Read.All'],
  probe: async (graph) => { await graph.get('/me/people?$top=1&$select=id'); },
  settingsSchema: [
    { key: 'maxPeople', label: 'Number of people', type: 'number', defaultValue: 8, min: 3, max: 30 },
    { key: 'showPresence', label: 'Show Teams presence', type: 'boolean', defaultValue: true }
  ],
  component: PeopleWidget
};

export default PeopleWidget;
