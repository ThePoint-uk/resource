import * as React from 'react';
import { Icon, IconButton } from '@fluentui/react';
import type { IWidgetDefinition, IWidgetProps } from '../../models';
import { useAsync } from '../../hooks/useAsync';
import { LoadingState, ErrorState, EmptyState, SectionHeader, FooterLink } from '../../components/shared/States';
import { ItemList, ItemRow } from '../../components/shared/ItemRow';
import shared from '../../components/shared/Shared.module.scss';
import type { IGraphEvent } from '../../services/GraphService';
import { formatDate, formatTime } from '../../utils/format';

function startOfDay(d: Date): Date {
  const r: Date = new Date(d);
  r.setHours(0, 0, 0, 0);
  return r;
}

function addDays(d: Date, days: number): Date {
  const r: Date = new Date(d);
  r.setDate(r.getDate() + days);
  return r;
}

function dayKey(d: Date): string {
  return `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`;
}

function dayLabel(d: Date): string {
  const today: Date = startOfDay(new Date());
  const diff: number = Math.round((startOfDay(d).getTime() - today.getTime()) / 86400000);
  if (diff === 0) {
    return 'Today';
  }
  if (diff === 1) {
    return 'Tomorrow';
  }
  return formatDate(d, { weekday: 'long', day: 'numeric', month: 'short' });
}

const CalendarWidget: React.FC<IWidgetProps> = ({ graph, settings, density }) => {
  const days: number = Number(settings.days) || 7;
  const max: number = Number(settings.maxItems) || 15;
  const [offset, setOffset] = React.useState<number>(0);
  const rangeStart: Date = addDays(startOfDay(new Date()), offset * days);
  const rangeEnd: Date = addDays(rangeStart, days);

  const events = useAsync<IGraphEvent[]>(() => graph.getCalendarView(rangeStart, rangeEnd, max), [graph, rangeStart.getTime(), rangeEnd.getTime(), max]);

  const grouped: { key: string; date: Date; events: IGraphEvent[] }[] = React.useMemo(() => {
    const map: Map<string, { key: string; date: Date; events: IGraphEvent[] }> = new Map();
    (events.data || []).forEach(e => {
      const start: Date = new Date(e.start.dateTime + (e.start.dateTime.endsWith('Z') ? '' : 'Z'));
      // Graph returns local-time strings when Prefer: outlook.timezone is set; parse accordingly.
      const local: Date = e.isAllDay ? new Date(e.start.dateTime.slice(0, 10) + 'T00:00:00') : new Date(e.start.dateTime.replace(/\.\d+$/, ''));
      const date: Date = isNaN(local.getTime()) ? start : local;
      const key: string = dayKey(date);
      if (!map.has(key)) {
        map.set(key, { key, date: startOfDay(date), events: [] });
      }
      map.get(key)!.events.push(e);
    });
    return Array.from(map.values()).sort((a, b) => a.date.getTime() - b.date.getTime());
  }, [events.data]);

  const now: Date = new Date();

  return (
    <div>
      <div className={shared.toolbar} style={{ justifyContent: 'space-between' }}>
        <span style={{ fontSize: 12.5, fontWeight: 600 }}>
          {formatDate(rangeStart, { day: 'numeric', month: 'short' })} – {formatDate(addDays(rangeEnd, -1), { day: 'numeric', month: 'short' })}
        </span>
        <span>
          <IconButton iconProps={{ iconName: 'ChevronLeft' }} title="Previous" onClick={() => setOffset(o => o - 1)} styles={{ root: { height: 26, width: 26 } }} />
          <IconButton iconProps={{ iconName: 'GotoToday' }} title="Today" onClick={() => setOffset(0)} disabled={offset === 0} styles={{ root: { height: 26, width: 26 } }} />
          <IconButton iconProps={{ iconName: 'ChevronRight' }} title="Next" onClick={() => setOffset(o => o + 1)} styles={{ root: { height: 26, width: 26 } }} />
        </span>
      </div>
      {events.loading && !events.data && <LoadingState rows={4} />}
      {events.error && <ErrorState message={events.error} onRetry={events.reload} />}
      {events.data && grouped.length === 0 && <EmptyState icon="Calendar" title="No events" message="Your calendar is clear for this period." />}
      {grouped.map(group => (
        <div key={group.key}>
          <SectionHeader text={dayLabel(group.date)} />
          <ItemList compact={density === 'compact'}>
            {group.events.map(e => {
              const start: Date = new Date(e.start.dateTime.replace(/\.\d+$/, ''));
              const end: Date = new Date(e.end.dateTime.replace(/\.\d+$/, ''));
              const inProgress: boolean = !e.isAllDay && start <= now && end >= now;
              const time: string = e.isAllDay ? 'All day' : `${formatTime(start)} – ${formatTime(end)}`;
              const joinUrl: string | undefined = e.onlineMeeting?.joinUrl;
              return (
                <ItemRow
                  key={e.id}
                  title={e.subject || '(No subject)'}
                  href={e.webLink}
                  meta={[time, e.location?.displayName, e.organizer?.emailAddress?.name].filter(Boolean).join(' · ')}
                  leading={
                    <span className={shared.rowIcon} style={{ background: inProgress ? 'var(--tpms-accent)' : 'var(--tpms-accent-soft)', color: inProgress ? '#fff' : 'var(--tpms-accent)', flexDirection: 'column', fontSize: 10, lineHeight: 1.1, fontWeight: 700 }}>
                      {e.isAllDay ? <Icon iconName="Calendar" /> : (<><span>{start.toLocaleTimeString(undefined, { hour: 'numeric' }).replace(/\s?[ap]m/i, '')}</span><span style={{ fontWeight: 400, fontSize: 9 }}>{start.getHours() < 12 ? 'AM' : 'PM'}</span></>)}
                    </span>
                  }
                  trailing={
                    joinUrl ? (
                      <a href={joinUrl} target="_blank" rel="noreferrer" className={`${shared.pill} ${shared.pillAccent}`} onClick={ev => ev.stopPropagation()} title="Join online meeting">
                        <Icon iconName="Video" /> Join
                      </a>
                    ) : e.responseStatus?.response === 'tentativelyAccepted' ? <span className={shared.pill}>Tentative</span> : undefined
                  }
                />
              );
            })}
          </ItemList>
        </div>
      ))}
      <FooterLink href="https://outlook.office.com/calendar/view/week" text="Open calendar" />
    </div>
  );
};

export const calendarDefinition: IWidgetDefinition = {
  type: 'calendar',
  title: 'My calendar',
  description: 'Upcoming meetings and events with one-click join for Teams meetings.',
  icon: 'Calendar',
  category: 'productivity',
  defaultSize: 1,
  allowMultiple: false,
  permissions: ['Calendars.Read'],
  probe: async (graph) => { await graph.get('/me/calendar?$select=id'); },
  settingsSchema: [
    { key: 'days', label: 'Days to show', type: 'number', defaultValue: 7, min: 1, max: 31 },
    { key: 'maxItems', label: 'Maximum events', type: 'number', defaultValue: 15, min: 3, max: 50 }
  ],
  component: CalendarWidget
};

export default CalendarWidget;
