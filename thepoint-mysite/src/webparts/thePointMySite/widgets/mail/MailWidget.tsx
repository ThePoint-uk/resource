import * as React from 'react';
import { Icon, Toggle } from '@fluentui/react';
import type { IWidgetDefinition, IWidgetProps } from '../../models';
import { useAsync } from '../../hooks/useAsync';
import { LoadingState, ErrorState, EmptyState, FooterLink } from '../../components/shared/States';
import { ItemList, ItemRow } from '../../components/shared/ItemRow';
import { Avatar } from '../../components/shared/Avatar';
import shared from '../../components/shared/Shared.module.scss';
import type { IGraphMessage } from '../../services/GraphService';
import { formatRelative, truncate } from '../../utils/format';

const MailWidget: React.FC<IWidgetProps> = ({ graph, settings, density, updateSettings }) => {
  const max: number = Number(settings.maxItems) || 8;
  const unreadOnly: boolean = settings.unreadOnly !== false;
  const messages = useAsync<IGraphMessage[]>(() => graph.getInboxMessages(max, unreadOnly), [graph, max, unreadOnly]);
  const counts = useAsync<{ unreadItemCount: number; totalItemCount: number }>(() => graph.getInboxUnreadCount(), [graph, messages.data]);

  return (
    <div>
      <div className={shared.toolbar} style={{ justifyContent: 'space-between' }}>
        <span className={`${shared.pill} ${shared.pillAccent}`}>
          <Icon iconName="Mail" /> {counts.data ? `${counts.data.unreadItemCount} unread` : '…'}
        </span>
        <Toggle inlineLabel label="Unread only" checked={unreadOnly} onChange={(_, v) => updateSettings({ unreadOnly: !!v })} styles={{ root: { margin: 0 }, label: { fontSize: 12 } }} />
      </div>
      {messages.loading && !messages.data && <LoadingState rows={4} />}
      {messages.error && <ErrorState message={messages.error} onRetry={messages.reload} />}
      {messages.data && messages.data.length === 0 && <EmptyState icon="MailCheck" title={unreadOnly ? 'Inbox zero' : 'No messages'} message={unreadOnly ? 'You have no unread mail.' : undefined} />}
      {messages.data && messages.data.length > 0 && (
        <ItemList compact={density === 'compact'}>
          {messages.data.map(m => (
            <ItemRow
              key={m.id}
              title={m.subject || '(No subject)'}
              href={m.webLink}
              leading={<Avatar graph={graph} name={m.from?.emailAddress?.name} size="sm" />}
              meta={
                <span style={{ fontWeight: m.isRead ? 400 : 600 }}>
                  {m.from?.emailAddress?.name || m.from?.emailAddress?.address}
                  {settings.showPreview !== false && m.bodyPreview ? ` — ${truncate(m.bodyPreview, 80)}` : ''}
                </span>
              }
              trailing={
                <>
                  {m.hasAttachments && <Icon iconName="Attach" />}
                  {m.importance === 'high' && <Icon iconName="Important" styles={{ root: { color: 'var(--tpms-danger)' } }} />}
                  <span>{formatRelative(m.receivedDateTime)}</span>
                </>
              }
            />
          ))}
        </ItemList>
      )}
      <FooterLink href="https://outlook.office.com/mail/" text="Open Outlook" />
    </div>
  );
};

export const mailDefinition: IWidgetDefinition = {
  type: 'mail',
  title: 'My inbox',
  description: 'Latest messages in your inbox with an unread counter.',
  icon: 'Mail',
  category: 'productivity',
  defaultSize: 1,
  allowMultiple: false,
  permissions: ['Mail.Read'],
  settingsSchema: [
    { key: 'maxItems', label: 'Messages to show', type: 'number', defaultValue: 8, min: 3, max: 30 },
    { key: 'unreadOnly', label: 'Unread only', type: 'boolean', defaultValue: true },
    { key: 'showPreview', label: 'Show message preview', type: 'boolean', defaultValue: true }
  ],
  component: MailWidget
};

export default MailWidget;
