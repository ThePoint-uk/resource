import * as React from 'react';
import { Checkbox, Dropdown, Icon, IconButton, Pivot, PivotItem, type IDropdownOption } from '@fluentui/react';
import type { IWidgetDefinition, IWidgetProps } from '../../models';
import { useAsync } from '../../hooks/useAsync';
import { LoadingState, ErrorState, EmptyState, FooterLink } from '../../components/shared/States';
import { ItemList, ItemRow } from '../../components/shared/ItemRow';
import shared from '../../components/shared/Shared.module.scss';
import type { IGraphPlannerTask, IGraphTodoList, IGraphTodoTask } from '../../services/GraphService';
import { formatDate } from '../../utils/format';

function dueLabel(iso: string | undefined): { text: string; overdue: boolean } | undefined {
  if (!iso) {
    return undefined;
  }
  const due: Date = new Date(iso);
  const today: Date = new Date();
  today.setHours(0, 0, 0, 0);
  const diff: number = Math.round((due.getTime() - today.getTime()) / 86400000);
  if (diff < 0) {
    return { text: `Overdue · ${formatDate(due, { day: 'numeric', month: 'short' })}`, overdue: true };
  }
  if (diff === 0) {
    return { text: 'Due today', overdue: false };
  }
  if (diff === 1) {
    return { text: 'Due tomorrow', overdue: false };
  }
  return { text: `Due ${formatDate(due, { weekday: 'short', day: 'numeric', month: 'short' })}`, overdue: false };
}

const TodoView: React.FC<IWidgetProps> = ({ graph, settings, density, updateSettings }) => {
  const max: number = Number(settings.maxItems) || 10;
  const lists = useAsync<IGraphTodoList[]>(() => graph.getTodoLists(), [graph]);
  const preferred: string | undefined = settings.todoListId ? String(settings.todoListId) : undefined;
  const listId: string | undefined = preferred && lists.data?.some(l => l.id === preferred) ? preferred : lists.data?.find(l => l.wellknownListName === 'defaultList')?.id || lists.data?.[0]?.id;
  const tasks = useAsync<IGraphTodoTask[]>(async () => (listId ? graph.getTodoTasks(listId, max) : []), [graph, listId, max]);
  const [newTitle, setNewTitle] = React.useState<string>('');
  const [busy, setBusy] = React.useState<string | undefined>(undefined);

  const complete = async (task: IGraphTodoTask): Promise<void> => {
    if (!listId) {
      return;
    }
    setBusy(task.id);
    try {
      await graph.completeTodoTask(listId, task.id);
      tasks.setData(prev => (prev || []).filter(t => t.id !== task.id));
    } finally {
      setBusy(undefined);
    }
  };

  const add = async (): Promise<void> => {
    if (!listId || !newTitle.trim()) {
      return;
    }
    setBusy('new');
    try {
      const created: IGraphTodoTask = await graph.createTodoTask(listId, newTitle.trim());
      tasks.setData(prev => [created, ...(prev || [])]);
      setNewTitle('');
    } finally {
      setBusy(undefined);
    }
  };

  if (lists.loading && !lists.data) {
    return <LoadingState rows={3} />;
  }
  if (lists.error) {
    return <ErrorState message={lists.error} onRetry={lists.reload} />;
  }
  const options: IDropdownOption[] = (lists.data || []).map(l => ({ key: l.id, text: l.displayName }));

  return (
    <div>
      {settings.showListPicker !== false && options.length > 1 && (
        <Dropdown options={options} selectedKey={listId} onChange={(_, o) => o && updateSettings({ todoListId: String(o.key) })} styles={{ root: { marginBottom: 6 }, title: { height: 30, lineHeight: 28, fontSize: 12.5 }, caretDownWrapper: { lineHeight: 28 } }} />
      )}
      {settings.allowAdd !== false && (
        <form
          className={shared.inlineForm}
          onSubmit={e => {
            e.preventDefault();
            add().catch(() => undefined);
          }}
        >
          <input value={newTitle} onChange={e => setNewTitle(e.target.value)} placeholder="Add a task" aria-label="New task title" disabled={busy === 'new'} />
          <IconButton iconProps={{ iconName: 'Add' }} title="Add task" type="submit" disabled={!newTitle.trim() || busy === 'new'} />
        </form>
      )}
      {tasks.loading && !tasks.data && <LoadingState rows={3} />}
      {tasks.error && <ErrorState message={tasks.error} onRetry={tasks.reload} />}
      {tasks.data && tasks.data.length === 0 && <EmptyState icon="CheckList" title="All caught up" message="No open tasks in this list." />}
      {tasks.data && tasks.data.length > 0 && (
        <ItemList compact={density === 'compact'}>
          {tasks.data.map(t => {
            const due = dueLabel(t.dueDateTime?.dateTime);
            return (
              <ItemRow
                key={t.id}
                title={t.title}
                meta={due ? <span style={{ color: due.overdue ? 'var(--tpms-danger)' : undefined }}>{due.text}</span> : undefined}
                leading={<Checkbox checked={false} disabled={busy === t.id} onChange={() => { complete(t).catch(() => undefined); }} ariaLabel={`Complete ${t.title}`} styles={{ root: { marginLeft: 4 } }} />}
                trailing={t.importance === 'high' ? <Icon iconName="Important" styles={{ root: { color: 'var(--tpms-danger)' } }} title="High importance" /> : undefined}
                href="https://to-do.office.com/tasks/"
              />
            );
          })}
        </ItemList>
      )}
    </div>
  );
};

const PlannerView: React.FC<IWidgetProps> = ({ graph, settings, density }) => {
  const max: number = Number(settings.maxItems) || 10;
  const tasks = useAsync<IGraphPlannerTask[]>(async () => {
    const all: IGraphPlannerTask[] = await graph.getPlannerTasks(200);
    const open: IGraphPlannerTask[] = all.filter(t => t.percentComplete < 100).sort((a, b) => (a.dueDateTime || '9999').localeCompare(b.dueDateTime || '9999')).slice(0, max);
    const planIds: string[] = open.map(t => t.planId || '').filter((v, i, arr) => v && arr.indexOf(v) === i);
    const titles: Map<string, string> = new Map();
    await Promise.all(
      planIds.map(async id => {
        try {
          const plan = await graph.getPlan(id);
          titles.set(id, plan.title);
        } catch {
          /* plan may be inaccessible */
        }
      })
    );
    return open.map(t => ({ ...t, planTitle: t.planId ? titles.get(t.planId) : undefined }));
  }, [graph, max]);

  if (tasks.loading && !tasks.data) {
    return <LoadingState rows={3} />;
  }
  if (tasks.error) {
    return <ErrorState message={tasks.error} onRetry={tasks.reload} />;
  }
  if (!tasks.data || tasks.data.length === 0) {
    return <EmptyState icon="PlannerLogo" title="No Planner tasks" message="Tasks assigned to you in Planner will show up here." />;
  }
  return (
    <ItemList compact={density === 'compact'}>
      {tasks.data.map(t => {
        const due = dueLabel(t.dueDateTime);
        return (
          <ItemRow
            key={t.id}
            title={t.title}
            href={`https://tasks.office.com/${graph.tenantId}/Home/Task/${t.id}`}
            icon={t.percentComplete > 0 ? 'ProgressRingDots' : 'CircleRing'}
            meta={[t.planTitle, due?.text].filter(Boolean).join(' · ')}
            trailing={t.priority !== undefined && t.priority <= 1 ? <Icon iconName="Important" styles={{ root: { color: 'var(--tpms-danger)' } }} title="Urgent" /> : undefined}
          />
        );
      })}
    </ItemList>
  );
};

const TasksWidget: React.FC<IWidgetProps> = (props) => {
  const { settings } = props;
  const [view, setView] = React.useState<string>(String(settings.defaultView || 'todo'));
  const showTodo: boolean = settings.showTodo !== false;
  const showPlanner: boolean = settings.showPlanner !== false;
  const effective: string = showTodo && showPlanner ? view : showTodo ? 'todo' : 'planner';
  return (
    <div>
      {showTodo && showPlanner && (
        <Pivot selectedKey={effective} onLinkClick={i => setView(i?.props.itemKey || 'todo')} headersOnly styles={{ root: { marginBottom: 6 }, link: { height: 30, fontSize: 12 }, linkIsSelected: { height: 30, fontSize: 12 } }}>
          <PivotItem headerText="To Do" itemKey="todo" />
          <PivotItem headerText="Planner" itemKey="planner" />
        </Pivot>
      )}
      {effective === 'todo' ? <TodoView {...props} /> : <PlannerView {...props} />}
      <FooterLink href={effective === 'todo' ? 'https://to-do.office.com/' : 'https://planner.cloud.microsoft/'} text={effective === 'todo' ? 'Open To Do' : 'Open Planner'} />
    </div>
  );
};

export const tasksDefinition: IWidgetDefinition = {
  type: 'tasks',
  title: 'My tasks',
  description: 'Open Microsoft To Do tasks (add and complete in place) and Planner tasks assigned to you.',
  icon: 'CheckList',
  category: 'productivity',
  defaultSize: 1,
  allowMultiple: false,
  permissions: ['Tasks.ReadWrite'],
  probe: async (graph) => { await graph.get('/me/todo/lists?$top=1&$select=id'); },
  settingsSchema: [
    { key: 'showTodo', label: 'Show To Do', type: 'boolean', defaultValue: true },
    { key: 'showPlanner', label: 'Show Planner', type: 'boolean', defaultValue: true },
    { key: 'defaultView', label: 'Default view', type: 'choice', defaultValue: 'todo', choices: [{ key: 'todo', text: 'To Do' }, { key: 'planner', text: 'Planner' }] },
    { key: 'allowAdd', label: 'Allow adding To Do tasks', type: 'boolean', defaultValue: true },
    { key: 'showListPicker', label: 'Show To Do list picker', type: 'boolean', defaultValue: true },
    { key: 'maxItems', label: 'Maximum tasks', type: 'number', defaultValue: 10, min: 3, max: 50 }
  ],
  component: TasksWidget
};

export default TasksWidget;
