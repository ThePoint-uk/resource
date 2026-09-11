import * as React from 'react';
import { IconButton, Spinner, SpinnerSize } from '@fluentui/react';
import type { IWidgetDefinition, IWidgetProps } from '../../models';
import { useAsync } from '../../hooks/useAsync';
import { LoadingState, ErrorState, EmptyState } from '../../components/shared/States';
import { Avatar } from '../../components/shared/Avatar';
import type { GraphService, IGraphUser } from '../../services/GraphService';
import { chatUrl } from '../team/PersonRow';
import styles from './MyOrgWidget.module.scss';

interface IOrgData {
  chain: IGraphUser[];
  me: IGraphUser;
  reports: IGraphUser[];
}

interface INodeProps {
  graph: GraphService;
  user: IGraphUser;
  depth: number;
  isMe?: boolean;
  initialChildren?: IGraphUser[];
  expandable: boolean;
  defaultExpanded?: boolean;
  maxDepth: number;
}

const OrgNode: React.FC<INodeProps> = ({ graph, user, depth, isMe, initialChildren, expandable, defaultExpanded, maxDepth }) => {
  const [expanded, setExpanded] = React.useState<boolean>(!!defaultExpanded);
  const [children, setChildren] = React.useState<IGraphUser[] | undefined>(initialChildren);
  const [loading, setLoading] = React.useState<boolean>(false);

  const toggle = (): void => {
    if (!expanded && children === undefined) {
      setLoading(true);
      graph
        .getDirectReports(user.id)
        .then(result => {
          setChildren(result);
          setLoading(false);
        })
        .catch(() => {
          setChildren([]);
          setLoading(false);
        });
    }
    setExpanded(!expanded);
  };

  const canExpand: boolean = expandable && depth < maxDepth && (children === undefined || children.length > 0);

  return (
    <div>
      <div className={`${styles.node} ${isMe ? styles.me : ''}`}>
        <Avatar graph={graph} userId={user.id} name={user.displayName} size="sm" />
        <div className={styles.body}>
          <div className={styles.name}>{user.displayName}{isMe ? ' (you)' : ''}</div>
          <div className={styles.title}>{[user.jobTitle, user.department].filter(Boolean).join(' · ')}</div>
        </div>
        {children && children.length > 0 && <span className={styles.count}>{children.length}</span>}
        <IconButton iconProps={{ iconName: 'TeamsLogo' }} title="Chat in Teams" href={chatUrl(user)} target="_blank" styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 13 } }} />
        {canExpand && (
          <span className={styles.expand}>
            {loading ? (
              <Spinner size={SpinnerSize.xSmall} />
            ) : (
              <IconButton iconProps={{ iconName: expanded ? 'ChevronUp' : 'ChevronDown' }} title={expanded ? 'Collapse' : 'Show reports'} onClick={toggle} styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 11 } }} />
            )}
          </span>
        )}
      </div>
      {expanded && children && children.length > 0 && (
        <div className={styles.children}>
          {children.map(c => (
            <OrgNode key={c.id} graph={graph} user={c} depth={depth + 1} expandable={expandable} maxDepth={maxDepth} />
          ))}
        </div>
      )}
    </div>
  );
};

const MyOrgWidget: React.FC<IWidgetProps> = ({ graph, settings }) => {
  const maxDepth: number = Number(settings.maxDepth) || 3;
  const showChain: boolean = settings.showChain !== false;
  const data = useAsync<IOrgData>(async () => {
    const [me, chain, reports] = await Promise.all([graph.getMe(), showChain ? graph.getManagerChain() : Promise.resolve([]), graph.getDirectReports('me')]);
    return { me, chain, reports };
  }, [graph, showChain]);

  if (data.loading && !data.data) {
    return <LoadingState rows={4} />;
  }
  if (data.error || !data.data) {
    return <ErrorState message={data.error || 'Organisation data unavailable'} onRetry={data.reload} />;
  }
  const { chain, me, reports } = data.data;
  if (chain.length === 0 && reports.length === 0) {
    return <EmptyState icon="Org" title="No organisation data" message="No manager or direct reports have been recorded for you in the directory." />;
  }

  // Chain comes nearest-manager first; render top-down.
  const topDown: IGraphUser[] = [...chain].reverse();

  return (
    <div className={styles.tree}>
      {topDown.map((u, i) => (
        <div key={u.id} style={{ marginLeft: i * 14 }}>
          <OrgNode graph={graph} user={u} depth={0} expandable={settings.expandManagers !== false} maxDepth={maxDepth} />
        </div>
      ))}
      <div style={{ marginLeft: topDown.length * 14 }}>
        <OrgNode graph={graph} user={me} depth={0} isMe initialChildren={reports} expandable defaultExpanded={settings.expandReports !== false} maxDepth={maxDepth} />
      </div>
    </div>
  );
};

export const myOrgDefinition: IWidgetDefinition = {
  type: 'org',
  title: 'My organisation',
  description: 'Your reporting line from the top of the organisation down to your direct reports, expandable at every level.',
  icon: 'Org',
  category: 'people',
  defaultSize: 1,
  allowMultiple: false,
  permissions: ['User.Read.All'],
  probe: async (graph) => { await graph.get('/me/directReports?$select=id&$top=1', { cacheKey: 'me:hasReports' }); },
  settingsSchema: [
    { key: 'showChain', label: 'Show management chain', type: 'boolean', defaultValue: true },
    { key: 'expandManagers', label: 'Allow expanding managers', type: 'boolean', defaultValue: true },
    { key: 'expandReports', label: 'Expand my reports by default', type: 'boolean', defaultValue: true },
    { key: 'maxDepth', label: 'Maximum expansion depth', type: 'number', defaultValue: 3, min: 1, max: 6 }
  ],
  component: MyOrgWidget
};

export default MyOrgWidget;
