import * as React from 'react';
import { DefaultButton, Icon, IconButton, MessageBar, MessageBarType, Panel, PanelType, Pivot, PivotItem, PrimaryButton, TextField, Spinner, SpinnerSize } from '@fluentui/react';
import type { IWidgetDefinition, IWidgetProps } from '../../models';
import { useAsync } from '../../hooks/useAsync';
import { Avatar } from '../../components/shared/Avatar';
import { LoadingState, ErrorState, SectionHeader } from '../../components/shared/States';
import shared from '../../components/shared/Shared.module.scss';
import styles from './MyProfileWidget.module.scss';
import type { IGraphPresence, IGraphUser } from '../../services/GraphService';
import { describeError } from '../../services/GraphService';

interface IProfileField {
  key: keyof IGraphUser;
  label: string;
  icon: string;
  kind: 'text' | 'multiline' | 'list' | 'phone' | 'mail' | 'url';
  group: 'contact' | 'work' | 'about';
}

const FIELDS: IProfileField[] = [
  { key: 'mail', label: 'Email', icon: 'Mail', kind: 'mail', group: 'contact' },
  { key: 'businessPhones', label: 'Business phone', icon: 'Phone', kind: 'phone', group: 'contact' },
  { key: 'mobilePhone', label: 'Mobile', icon: 'CellPhone', kind: 'phone', group: 'contact' },
  { key: 'officeLocation', label: 'Office', icon: 'POI', kind: 'text', group: 'contact' },
  { key: 'streetAddress', label: 'Street', icon: 'MapPin', kind: 'text', group: 'contact' },
  { key: 'city', label: 'City', icon: 'CityNext', kind: 'text', group: 'contact' },
  { key: 'state', label: 'State / county', icon: 'Nav2DMapView', kind: 'text', group: 'contact' },
  { key: 'postalCode', label: 'Postcode', icon: 'NumberField', kind: 'text', group: 'contact' },
  { key: 'country', label: 'Country', icon: 'World', kind: 'text', group: 'contact' },
  { key: 'jobTitle', label: 'Job title', icon: 'Work', kind: 'text', group: 'work' },
  { key: 'department', label: 'Department', icon: 'Org', kind: 'text', group: 'work' },
  { key: 'companyName', label: 'Company', icon: 'CityNext2', kind: 'text', group: 'work' },
  { key: 'employeeId', label: 'Employee ID', icon: 'ContactCard', kind: 'text', group: 'work' },
  { key: 'userPrincipalName', label: 'Sign-in name', icon: 'Signin', kind: 'text', group: 'work' },
  { key: 'preferredLanguage', label: 'Language', icon: 'LocaleLanguage', kind: 'text', group: 'work' },
  { key: 'hireDate', label: 'Hire date', icon: 'Calendar', kind: 'text', group: 'work' },
  { key: 'preferredName', label: 'Preferred name', icon: 'Contact', kind: 'text', group: 'about' },
  { key: 'aboutMe', label: 'About me', icon: 'Info', kind: 'multiline', group: 'about' },
  { key: 'skills', label: 'Skills', icon: 'Lightbulb', kind: 'list', group: 'about' },
  { key: 'interests', label: 'Interests', icon: 'Heart', kind: 'list', group: 'about' },
  { key: 'schools', label: 'Schools', icon: 'Education', kind: 'list', group: 'about' },
  { key: 'pastProjects', label: 'Past projects', icon: 'ProjectCollection', kind: 'list', group: 'about' },
  { key: 'responsibilities', label: 'Responsibilities', icon: 'CheckList', kind: 'list', group: 'about' },
  { key: 'mySite', label: 'My site', icon: 'Link', kind: 'url', group: 'about' }
];

/** Fields Microsoft Graph lets a standard user update on their own object with User.ReadWrite. */
export const SELF_EDITABLE_DEFAULT: string = 'mobilePhone,businessPhones,preferredName,aboutMe,skills,interests,schools,pastProjects,responsibilities,mySite';

function toDisplay(value: unknown, kind: IProfileField['kind']): string {
  if (value === undefined || value === null || value === '') {
    return '';
  }
  if (Array.isArray(value)) {
    return value.join(', ');
  }
  if (kind === 'text' && typeof value === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(value)) {
    return new Date(value).toLocaleDateString();
  }
  return String(value);
}

function toEditable(value: unknown): string {
  if (Array.isArray(value)) {
    return value.join(', ');
  }
  return value === undefined || value === null ? '' : String(value);
}

function fromEditable(value: string, field: IProfileField): unknown {
  if (field.kind === 'list') {
    return value.split(',').map(v => v.trim()).filter(v => v.length > 0);
  }
  if (field.key === 'businessPhones') {
    return value.split(',').map(v => v.trim()).filter(v => v.length > 0);
  }
  return value.trim();
}

const MyProfileWidget: React.FC<IWidgetProps> = (props) => {
  const { graph, sp, settings, size, updateSettings } = props;
  const editable: string[] = String(settings.editableFields || SELF_EDITABLE_DEFAULT).split(',').map(s => s.trim()).filter(Boolean);
  const showAbout: boolean = settings.showAbout !== false;
  const showWork: boolean = settings.showWork !== false;
  const showPresence: boolean = settings.showPresence !== false;

  const me = useAsync<IGraphUser>(() => graph.getMe(), [graph]);
  const presence = useAsync<IGraphPresence | undefined>(
    async () => (showPresence ? graph.get<IGraphPresence>('/me/presence', { cacheKey: 'me:presence', noCache: true }) : undefined),
    [graph, showPresence]
  );

  const [editing, setEditing] = React.useState<boolean>(false);
  const [draft, setDraft] = React.useState<Record<string, string>>({});
  const [saving, setSaving] = React.useState<boolean>(false);
  const [saveMessage, setSaveMessage] = React.useState<{ type: MessageBarType; text: string } | undefined>(undefined);
  const [tab, setTab] = React.useState<string>(String(settings.defaultTab || 'contact'));

  const openEditor = (): void => {
    const next: Record<string, string> = {};
    FIELDS.filter(f => editable.indexOf(f.key) >= 0).forEach(f => {
      next[f.key] = toEditable(me.data?.[f.key]);
    });
    setDraft(next);
    setSaveMessage(undefined);
    setEditing(true);
  };

  const save = async (): Promise<void> => {
    if (!me.data) {
      return;
    }
    setSaving(true);
    setSaveMessage(undefined);
    const changed: Partial<Record<string, unknown>> = {};
    FIELDS.filter(f => editable.indexOf(f.key) >= 0).forEach(f => {
      const before: string = toEditable(me.data?.[f.key]);
      if (draft[f.key] !== undefined && draft[f.key] !== before) {
        changed[f.key] = fromEditable(draft[f.key], f);
      }
    });
    const keys: string[] = Object.keys(changed);
    if (keys.length === 0) {
      setSaving(false);
      setEditing(false);
      return;
    }
    try {
      await graph.updateMe(changed);
      setSaveMessage({ type: MessageBarType.success, text: 'Profile updated. Changes can take a few minutes to appear everywhere.' });
      me.reload();
      setEditing(false);
    } catch (bulkError) {
      // Retry field by field so the user learns exactly which field is locked.
      const failed: string[] = [];
      for (const key of keys) {
        try {
          await graph.updateMe({ [key]: changed[key] });
        } catch {
          failed.push(FIELDS.filter(f => f.key === key)[0]?.label || key);
        }
      }
      if (failed.length === keys.length) {
        setSaveMessage({ type: MessageBarType.error, text: `Could not save: ${describeError(bulkError)}` });
      } else {
        setSaveMessage({
          type: MessageBarType.warning,
          text: `Saved, except: ${failed.join(', ')}. These fields are managed by your organisation and cannot be edited here.`
        });
        me.reload();
      }
    } finally {
      setSaving(false);
    }
  };

  if (me.loading && !me.data) {
    return <LoadingState rows={3} />;
  }
  if (me.error || !me.data) {
    return <ErrorState message={me.error || 'Profile unavailable'} onRetry={me.reload} />;
  }

  const user: IGraphUser = me.data;
  const renderFields = (group: IProfileField['group']): React.ReactNode => {
    const fields: IProfileField[] = FIELDS.filter(f => f.group === group).filter(f => toDisplay(user[f.key], f.kind) !== '' || editable.indexOf(f.key) >= 0);
    if (fields.length === 0) {
      return <div className={styles.editHelp}>Nothing to show yet.</div>;
    }
    return (
      <dl className={styles.fields}>
        {fields.map(f => {
          const value: string = toDisplay(user[f.key], f.kind);
          return (
            <div className={styles.field} key={f.key}>
              <Icon iconName={f.icon} className={styles.fieldIcon} />
              <div className={styles.fieldBody}>
                <dt className={styles.fieldLabel}>
                  {f.label}
                  {editable.indexOf(f.key) < 0 && <Icon iconName="Lock" className={styles.readonlyBadge} title="Managed by your organisation" />}
                </dt>
                <dd className={styles.fieldValue} style={{ margin: 0 }}>
                  {value === '' ? <span style={{ color: 'var(--tpms-text-muted)' }}>Not set</span> : renderValue(value, f)}
                </dd>
              </div>
            </div>
          );
        })}
      </dl>
    );
  };

  const canEdit: boolean = editable.length > 0 && settings.allowEdit !== false;
  const accountUrl: string = 'https://myaccount.microsoft.com/';
  const delveUrl: string = `${sp.tenantUrl}/_layouts/15/me.aspx`;

  return (
    <div className={size >= 2 ? styles.wide : ''}>
      <div className={styles.hero}>
        <Avatar graph={graph} userId="me" name={user.displayName} size={size >= 2 ? 'xl' : 'lg'} presence={presence.data?.availability} />
        <div className={styles.heroBody}>
          <div className={styles.name}>{user.preferredName || user.displayName}</div>
          <div className={styles.job}>{[user.jobTitle, user.department].filter(Boolean).join(' · ')}</div>
          {presence.data && (
            <div className={styles.job}>
              <span className={`${shared.pill} ${shared.pillAccent}`}>{presence.data.availability.replace(/([A-Z])/g, ' $1').trim()}</span>
            </div>
          )}
          <div className={styles.heroActions}>
            {canEdit && <PrimaryButton text="Edit profile" iconProps={{ iconName: 'Edit' }} onClick={openEditor} styles={{ root: { minWidth: 0, height: 28 }, label: { fontSize: 12 } }} />}
            <DefaultButton text="My account" href={accountUrl} target="_blank" iconProps={{ iconName: 'Settings' }} styles={{ root: { minWidth: 0, height: 28 }, label: { fontSize: 12 } }} />
            <IconButton iconProps={{ iconName: 'ContactCard' }} title="Open profile page" href={delveUrl} target="_blank" />
          </div>
        </div>
      </div>

      {saveMessage && (
        <MessageBar messageBarType={saveMessage.type} onDismiss={() => setSaveMessage(undefined)} styles={{ root: { marginBottom: 8 } }}>
          {saveMessage.text}
        </MessageBar>
      )}

      <Pivot
        selectedKey={tab}
        onLinkClick={item => {
          const key: string = item?.props.itemKey || 'contact';
          setTab(key);
          updateSettings({ defaultTab: key });
        }}
        styles={{ link: { height: 32, fontSize: 12.5 }, linkIsSelected: { height: 32, fontSize: 12.5 } }}
        headersOnly
      >
        <PivotItem headerText="Contact" itemKey="contact" />
        {showWork && <PivotItem headerText="Work" itemKey="work" />}
        {showAbout && <PivotItem headerText="About" itemKey="about" />}
      </Pivot>
      <div style={{ paddingTop: 8 }}>
        {tab === 'contact' && renderFields('contact')}
        {tab === 'work' && showWork && renderFields('work')}
        {tab === 'about' && showAbout && (
          <>
            {user.aboutMe && (
              <>
                <SectionHeader text="About me" />
                <div className={styles.about}>{user.aboutMe}</div>
              </>
            )}
            {renderFields('about')}
          </>
        )}
      </div>

      <Panel
        isOpen={editing}
        onDismiss={() => !saving && setEditing(false)}
        type={PanelType.medium}
        headerText="Edit my profile"
        closeButtonAriaLabel="Close"
        isFooterAtBottom
        onRenderFooterContent={() => (
          <div style={{ display: 'flex', gap: 8 }}>
            <PrimaryButton text={saving ? 'Saving…' : 'Save'} onClick={() => { save().catch(() => undefined); }} disabled={saving} />
            <DefaultButton text="Cancel" onClick={() => setEditing(false)} disabled={saving} />
            {saving && <Spinner size={SpinnerSize.small} />}
          </div>
        )}
      >
        <div className={styles.editGrid}>
          <div className={styles.editHelp}>
            Only fields your organisation allows you to manage are shown. Lists such as skills and interests are comma separated.
          </div>
          {saveMessage && <MessageBar messageBarType={saveMessage.type}>{saveMessage.text}</MessageBar>}
          {FIELDS.filter(f => editable.indexOf(f.key) >= 0).map(f => (
            <TextField
              key={f.key}
              label={f.label}
              multiline={f.kind === 'multiline'}
              rows={f.kind === 'multiline' ? 4 : undefined}
              value={draft[f.key] || ''}
              onChange={(_, v) => setDraft(prev => ({ ...prev, [f.key]: v || '' }))}
              description={f.kind === 'list' || f.key === 'businessPhones' ? 'Comma separated' : undefined}
              disabled={saving}
            />
          ))}
        </div>
      </Panel>
    </div>
  );
};

function renderValue(value: string, field: IProfileField): React.ReactNode {
  if (field.kind === 'mail') {
    return <a href={`mailto:${value}`}>{value}</a>;
  }
  if (field.kind === 'phone') {
    return value.split(',').map((p, i) => (
      <span key={i}>
        {i > 0 && ', '}
        <a href={`tel:${p.trim()}`}>{p.trim()}</a>
      </span>
    ));
  }
  if (field.kind === 'url') {
    return <a href={value} target="_blank" rel="noreferrer">{value}</a>;
  }
  if (field.kind === 'list') {
    return (
      <span className={shared.chips}>
        {value.split(',').map((v, i) => (
          <span className={shared.chip} key={i}>{v.trim()}</span>
        ))}
      </span>
    );
  }
  return value;
}

export const myProfileDefinition: IWidgetDefinition = {
  type: 'profile',
  title: 'My profile',
  description: 'Your photo, contact details and about-me information, with in-place editing of the fields you own.',
  icon: 'Contact',
  category: 'me',
  defaultSize: 1,
  allowMultiple: false,
  permissions: ['User.ReadWrite', 'Presence.Read'],
  probe: async (graph) => { await graph.get('/me?$select=id', { cacheKey: 'probe:me' }); },
  settingsSchema: [
    { key: 'allowEdit', label: 'Allow editing', type: 'boolean', defaultValue: true },
    { key: 'showPresence', label: 'Show Teams presence', type: 'boolean', defaultValue: true },
    { key: 'showWork', label: 'Show Work tab', type: 'boolean', defaultValue: true },
    { key: 'showAbout', label: 'Show About tab', type: 'boolean', defaultValue: true },
    {
      key: 'editableFields',
      label: 'Editable fields',
      type: 'text',
      defaultValue: SELF_EDITABLE_DEFAULT,
      description: 'Comma separated Graph user property names that users may edit. Directory-synced fields will be rejected by Microsoft Graph.'
    }
  ],
  component: MyProfileWidget
};

export default MyProfileWidget;
