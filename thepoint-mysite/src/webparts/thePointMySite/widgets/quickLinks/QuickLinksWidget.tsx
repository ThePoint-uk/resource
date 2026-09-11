import * as React from 'react';
import { DefaultButton, Icon, IconButton, PrimaryButton, TextField } from '@fluentui/react';
import type { IQuickLink, IWidgetDefinition, IWidgetProps } from '../../models';
import { EmptyState } from '../../components/shared/States';
import { ItemList, ItemRow } from '../../components/shared/ItemRow';
import shared from '../../components/shared/Shared.module.scss';
import { uid } from '../../utils/format';

function guessIcon(url: string): string {
  const u: string = url.toLowerCase();
  if (u.indexOf('sharepoint.com') >= 0) return 'SharepointLogo';
  if (u.indexOf('teams.microsoft') >= 0) return 'TeamsLogo';
  if (u.indexOf('outlook') >= 0) return 'OutlookLogo';
  if (u.indexOf('powerbi') >= 0) return 'PowerBILogo';
  if (u.indexOf('github') >= 0) return 'GitGraph';
  if (u.indexOf('linkedin') >= 0) return 'LinkedInLogo';
  if (u.indexOf('youtube') >= 0) return 'Video';
  return 'Link';
}

const QuickLinksWidget: React.FC<IWidgetProps> = ({ settings, updateSettings, editMode, density }) => {
  const links: IQuickLink[] = Array.isArray(settings.links) ? (settings.links as IQuickLink[]) : [];
  const layout: string = String(settings.layout || 'list');
  const [adding, setAdding] = React.useState<boolean>(false);
  const [editing, setEditing] = React.useState<IQuickLink | undefined>(undefined);
  const [title, setTitle] = React.useState<string>('');
  const [url, setUrl] = React.useState<string>('');
  const [icon, setIcon] = React.useState<string>('');

  const startAdd = (): void => {
    setEditing(undefined);
    setTitle('');
    setUrl('');
    setIcon('');
    setAdding(true);
  };
  const startEdit = (link: IQuickLink): void => {
    setEditing(link);
    setTitle(link.title);
    setUrl(link.url);
    setIcon(link.icon || '');
    setAdding(true);
  };
  const commit = (): void => {
    const cleanUrl: string = /^https?:\/\//i.test(url.trim()) ? url.trim() : `https://${url.trim()}`;
    const link: IQuickLink = { id: editing?.id || uid('link'), title: title.trim() || cleanUrl, url: cleanUrl, icon: icon.trim() || guessIcon(cleanUrl) };
    const next: IQuickLink[] = editing ? links.map(l => (l.id === editing.id ? link : l)) : [...links, link];
    updateSettings({ links: next });
    setAdding(false);
  };
  const remove = (id: string): void => updateSettings({ links: links.filter(l => l.id !== id) });
  const move = (id: string, dir: -1 | 1): void => {
    const idx: number = links.findIndex(l => l.id === id);
    const target: number = idx + dir;
    if (idx < 0 || target < 0 || target >= links.length) {
      return;
    }
    const next: IQuickLink[] = [...links];
    const [item] = next.splice(idx, 1);
    next.splice(target, 0, item);
    updateSettings({ links: next });
  };

  const manage: boolean = settings.allowUserLinks !== false;

  return (
    <div>
      {links.length === 0 && !adding && (
        <EmptyState icon="Link" title="No links yet" message={manage ? 'Add your favourite places to get to them in one click.' : undefined}>
          {manage && <PrimaryButton text="Add a link" iconProps={{ iconName: 'Add' }} onClick={startAdd} />}
        </EmptyState>
      )}
      {links.length > 0 && layout === 'tiles' && (
        <div className={shared.grid}>
          {links.map(l => (
            <a key={l.id} className={shared.tile} href={l.url} target="_blank" rel="noreferrer" title={l.url}>
              <span className={shared.tileIcon}><Icon iconName={l.icon || 'Link'} styles={{ root: { color: 'var(--tpms-accent)' } }} /></span>
              <span className={shared.tileTitle}>{l.title}</span>
            </a>
          ))}
        </div>
      )}
      {links.length > 0 && layout !== 'tiles' && (
        <ItemList compact={density === 'compact'}>
          {links.map(l => (
            <ItemRow
              key={l.id}
              title={l.title}
              href={l.url}
              icon={l.icon || 'Link'}
              iconColor="var(--tpms-accent)"
              meta={l.url.replace(/^https?:\/\//, '')}
              trailing={
                manage && editMode ? (
                  <>
                    <IconButton iconProps={{ iconName: 'Up' }} title="Move up" onClick={e => { e.preventDefault(); move(l.id, -1); }} styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 11 } }} />
                    <IconButton iconProps={{ iconName: 'Down' }} title="Move down" onClick={e => { e.preventDefault(); move(l.id, 1); }} styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 11 } }} />
                    <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={e => { e.preventDefault(); startEdit(l); }} styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 12 } }} />
                    <IconButton iconProps={{ iconName: 'Delete' }} title="Remove" onClick={e => { e.preventDefault(); remove(l.id); }} styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 12 } }} />
                  </>
                ) : undefined
              }
            />
          ))}
        </ItemList>
      )}
      {adding && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6, paddingTop: 8, borderTop: '1px solid var(--tpms-border)', marginTop: 8 }}>
          <TextField label="Title" value={title} onChange={(_, v) => setTitle(v || '')} />
          <TextField label="URL" value={url} onChange={(_, v) => setUrl(v || '')} placeholder="https://" />
          <TextField label="Icon (optional Fluent icon name)" value={icon} onChange={(_, v) => setIcon(v || '')} placeholder="e.g. Link, Globe, Document" />
          <div style={{ display: 'flex', gap: 6 }}>
            <PrimaryButton text={editing ? 'Save' : 'Add'} onClick={commit} disabled={!url.trim()} />
            <DefaultButton text="Cancel" onClick={() => setAdding(false)} />
          </div>
        </div>
      )}
      {manage && links.length > 0 && !adding && (
        <div style={{ paddingTop: 6 }}>
          <DefaultButton text="Add link" iconProps={{ iconName: 'Add' }} onClick={startAdd} styles={{ root: { height: 28, minWidth: 0 }, label: { fontSize: 12 } }} />
          {!editMode && <span style={{ fontSize: 11, color: 'var(--tpms-text-muted)', marginLeft: 8 }}>Use Customise to reorder or remove links.</span>}
        </div>
      )}
    </div>
  );
};

export const quickLinksDefinition: IWidgetDefinition = {
  type: 'quickLinks',
  title: 'Quick links',
  description: 'Your own personal bookmarks, shown as a list or tiles.',
  icon: 'Link',
  category: 'personal',
  defaultSize: 1,
  allowMultiple: true,
  settingsSchema: [
    { key: 'layout', label: 'Layout', type: 'choice', defaultValue: 'list', choices: [{ key: 'list', text: 'List' }, { key: 'tiles', text: 'Tiles' }] },
    { key: 'allowUserLinks', label: 'Allow users to manage links', type: 'boolean', defaultValue: true },
    { key: 'links', label: 'Links', type: 'links', defaultValue: [] }
  ],
  component: QuickLinksWidget
};

export default QuickLinksWidget;
