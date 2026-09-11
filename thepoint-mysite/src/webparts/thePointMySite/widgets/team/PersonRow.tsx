import * as React from 'react';
import { Icon, IconButton } from '@fluentui/react';
import shared from '../../components/shared/Shared.module.scss';
import { Avatar } from '../../components/shared/Avatar';
import type { GraphService, IGraphUser } from '../../services/GraphService';

export interface IPersonRowProps {
  graph: GraphService;
  user: IGraphUser;
  presence?: string;
  badge?: string;
  compact?: boolean;
  onClick?: () => void;
}

export function chatUrl(user: IGraphUser): string {
  const upn: string = user.userPrincipalName || user.mail || '';
  return `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(upn)}`;
}

export function callUrl(user: IGraphUser): string {
  const upn: string = user.userPrincipalName || user.mail || '';
  return `https://teams.microsoft.com/l/call/0/0?users=${encodeURIComponent(upn)}`;
}

/** A person row with avatar, name, title and quick chat / mail / call actions. */
export const PersonRow: React.FC<IPersonRowProps> = ({ graph, user, presence, badge, compact, onClick }) => {
  const [hover, setHover] = React.useState<boolean>(false);
  return (
    <div
      className={shared.row}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      onFocus={() => setHover(true)}
      onBlur={() => setHover(false)}
      onClick={onClick}
      role={onClick ? 'button' : undefined}
      tabIndex={onClick ? 0 : undefined}
      onKeyDown={e => {
        if (onClick && (e.key === 'Enter' || e.key === ' ')) {
          e.preventDefault();
          onClick();
        }
      }}
      style={{ cursor: onClick ? 'pointer' : 'default' }}
    >
      <Avatar graph={graph} userId={user.id} name={user.displayName} size={compact ? 'sm' : 'md'} presence={presence} />
      <span className={shared.rowBody}>
        <span className={shared.rowTitle}>
          {user.displayName}
          {badge && <span className={`${shared.pill} ${shared.pillAccent}`} style={{ marginLeft: 6 }}>{badge}</span>}
        </span>
        <span className={shared.rowMeta}>{[user.jobTitle, user.department].filter(Boolean).join(' · ') || user.mail}</span>
      </span>
      <span className={shared.rowTrailing} style={{ opacity: hover ? 1 : 0.55 }}>
        <IconButton iconProps={{ iconName: 'TeamsLogo' }} title="Chat in Teams" href={chatUrl(user)} target="_blank" styles={{ root: { height: 26, width: 26 }, icon: { fontSize: 14 } }} onClick={e => e.stopPropagation()} />
        {user.mail && <IconButton iconProps={{ iconName: 'Mail' }} title="Send email" href={`mailto:${user.mail}`} styles={{ root: { height: 26, width: 26 }, icon: { fontSize: 14 } }} onClick={e => e.stopPropagation()} />}
        {(user.businessPhones?.[0] || user.mobilePhone) && (
          <IconButton iconProps={{ iconName: 'Phone' }} title={user.mobilePhone || user.businessPhones?.[0]} href={`tel:${user.mobilePhone || user.businessPhones?.[0]}`} styles={{ root: { height: 26, width: 26 }, icon: { fontSize: 14 } }} onClick={e => e.stopPropagation()} />
        )}
        {onClick && <Icon iconName="ChevronRight" styles={{ root: { fontSize: 10 } }} />}
      </span>
    </div>
  );
};
