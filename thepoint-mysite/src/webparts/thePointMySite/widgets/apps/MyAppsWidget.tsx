import * as React from 'react';
import { Icon, SearchBox } from '@fluentui/react';
import type { IWidgetDefinition, IWidgetProps } from '../../models';
import { useAsync } from '../../hooks/useAsync';
import { LoadingState, ErrorState, EmptyState, SectionHeader, FooterLink } from '../../components/shared/States';
import { ItemList, ItemRow } from '../../components/shared/ItemRow';
import shared from '../../components/shared/Shared.module.scss';
import type { IGraphAppRoleAssignment, IGraphServicePrincipal } from '../../services/GraphService';

interface IApp {
  id: string;
  name: string;
  url: string;
  logoUrl?: string;
  icon?: string;
  color?: string;
  source: 'entra' | 'm365';
}

/** Well-known Microsoft 365 launch points. */
const M365_APPS: IApp[] = [
  { id: 'outlook', name: 'Outlook', url: 'https://outlook.office.com/mail/', icon: 'OutlookLogo', color: '#0078d4', source: 'm365' },
  { id: 'teams', name: 'Teams', url: 'https://teams.microsoft.com/', icon: 'TeamsLogo', color: '#6264a7', source: 'm365' },
  { id: 'onedrive', name: 'OneDrive', url: 'https://www.microsoft365.com/launch/onedrive', icon: 'OneDriveLogo', color: '#0364b8', source: 'm365' },
  { id: 'sharepoint', name: 'SharePoint', url: 'https://www.microsoft365.com/launch/sharepoint', icon: 'SharepointLogo', color: '#038387', source: 'm365' },
  { id: 'word', name: 'Word', url: 'https://www.microsoft365.com/launch/word', icon: 'WordLogo', color: '#185abd', source: 'm365' },
  { id: 'excel', name: 'Excel', url: 'https://www.microsoft365.com/launch/excel', icon: 'ExcelLogo', color: '#107c41', source: 'm365' },
  { id: 'powerpoint', name: 'PowerPoint', url: 'https://www.microsoft365.com/launch/powerpoint', icon: 'PowerPointLogo', color: '#c43e1c', source: 'm365' },
  { id: 'onenote', name: 'OneNote', url: 'https://www.microsoft365.com/launch/onenote', icon: 'OneNoteLogo', color: '#7719aa', source: 'm365' },
  { id: 'planner', name: 'Planner', url: 'https://planner.cloud.microsoft/', icon: 'PlannerLogo', color: '#31752f', source: 'm365' },
  { id: 'todo', name: 'To Do', url: 'https://to-do.office.com/', icon: 'CheckList', color: '#2564cf', source: 'm365' },
  { id: 'forms', name: 'Forms', url: 'https://forms.office.com/', icon: 'OfficeFormsLogo', color: '#008272', source: 'm365' },
  { id: 'powerbi', name: 'Power BI', url: 'https://app.powerbi.com/', icon: 'PowerBILogo', color: '#f2c811', source: 'm365' },
  { id: 'powerapps', name: 'Power Apps', url: 'https://make.powerapps.com/', icon: 'PowerApps', color: '#742774', source: 'm365' },
  { id: 'copilot', name: 'Copilot', url: 'https://m365.cloud.microsoft/chat/', icon: 'Robot', color: '#0f6cbd', source: 'm365' },
  { id: 'loop', name: 'Loop', url: 'https://loop.cloud.microsoft/', icon: 'Sync', color: '#8e44ad', source: 'm365' },
  { id: 'stream', name: 'Stream', url: 'https://www.microsoft365.com/launch/stream', icon: 'StreamLogo', color: '#bc1948', source: 'm365' },
  { id: 'bookings', name: 'Bookings', url: 'https://outlook.office.com/bookings/', icon: 'BookingsLogo', color: '#00a4ef', source: 'm365' },
  { id: 'viva', name: 'Viva Engage', url: 'https://engage.cloud.microsoft/', icon: 'YammerLogo', color: '#106ebe', source: 'm365' }
];

const MyAppsWidget: React.FC<IWidgetProps> = ({ graph, settings }) => {
  const showEntra: boolean = settings.showEntra !== false;
  const showM365: boolean = settings.showM365 !== false;
  const layout: string = String(settings.layout || 'tiles');
  const m365Keys: string[] = String(settings.m365Apps || M365_APPS.map(a => a.id).join(',')).split(',').map(s => s.trim()).filter(Boolean);
  const [query, setQuery] = React.useState<string>('');
  const tenantId: string = graph.tenantId;

  const entra = useAsync<{ apps: IApp[]; degraded: boolean }>(async () => {
    if (!showEntra) {
      return { apps: [], degraded: false };
    }
    const assignments: IGraphAppRoleAssignment[] = await graph.getAppRoleAssignments();
    const seen: Set<string> = new Set();
    const uniqueAssignments: IGraphAppRoleAssignment[] = assignments.filter(a => {
      if (seen.has(a.resourceId)) {
        return false;
      }
      seen.add(a.resourceId);
      return true;
    });
    let sps: IGraphServicePrincipal[] = [];
    let degraded: boolean = false;
    try {
      sps = await graph.getServicePrincipals(uniqueAssignments.map(a => a.resourceId));
    } catch {
      degraded = true;
    }
    const spMap: Map<string, IGraphServicePrincipal> = new Map();
    sps.forEach(s => spMap.set(s.id, s));
    const apps: IApp[] = uniqueAssignments
      .map((a): IApp | undefined => {
        const spn: IGraphServicePrincipal | undefined = spMap.get(a.resourceId);
        if (spn && spn.tags && spn.tags.indexOf('HideApp') >= 0) {
          return undefined;
        }
        if (spn && spn.accountEnabled === false) {
          return undefined;
        }
        // Hide Microsoft first-party service principals that are not launchable apps.
        if (spn && !spn.homepage && !spn.loginUrl && !(spn.tags || []).some(t => t.indexOf('WindowsAzureActiveDirectoryIntegratedApp') >= 0 || t.indexOf('WindowsAzureActiveDirectoryGalleryApplication') >= 0 || t.indexOf('WindowsAzureActiveDirectoryCustomSingleSignOnApplication') >= 0)) {
          return undefined;
        }
        const url: string = spn
          ? `https://launcher.myapps.microsoft.com/api/signin/${spn.appId}?tenantId=${tenantId}`
          : `https://myapplications.microsoft.com/?tenantId=${tenantId}`;
        return {
          id: a.resourceId,
          name: spn?.displayName || a.resourceDisplayName,
          url,
          logoUrl: spn?.info?.logoUrl,
          icon: 'AppIconDefault',
          source: 'entra'
        };
      })
      .filter((a): a is IApp => !!a)
      .sort((a, b) => a.name.localeCompare(b.name));
    return { apps, degraded };
  }, [graph, showEntra, tenantId]);

  const m365: IApp[] = showM365 ? M365_APPS.filter(a => m365Keys.indexOf(a.id) >= 0) : [];
  const q: string = query.trim().toLowerCase();
  const filter = (apps: IApp[]): IApp[] => (q ? apps.filter(a => a.name.toLowerCase().indexOf(q) >= 0) : apps);
  const entraApps: IApp[] = filter(entra.data?.apps || []);
  const m365Apps: IApp[] = filter(m365);
  const myAppsUrl: string = `https://myapplications.microsoft.com/?tenantId=${tenantId}`;

  const renderApps = (apps: IApp[]): React.ReactNode =>
    layout === 'list' ? (
      <ItemList>
        {apps.map(a => (
          <ItemRow key={a.id} title={a.name} href={a.url} imageUrl={a.logoUrl} icon={a.icon} iconColor={a.color} meta={a.source === 'entra' ? 'Entra application' : 'Microsoft 365'} />
        ))}
      </ItemList>
    ) : (
      <div className={shared.grid}>
        {apps.map(a => (
          <a key={a.id} className={shared.tile} href={a.url} target="_blank" rel="noreferrer" title={a.name}>
            <span className={shared.tileIcon}>{a.logoUrl ? <img src={a.logoUrl} alt="" loading="lazy" /> : <Icon iconName={a.icon || 'AppIconDefault'} styles={{ root: { color: a.color } }} />}</span>
            <span className={shared.tileTitle}>{a.name}</span>
          </a>
        ))}
      </div>
    );

  return (
    <div>
      {settings.showSearch !== false && (
        <div className={shared.toolbar}>
          <SearchBox placeholder="Find an app" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { flex: 1, height: 30 } }} />
        </div>
      )}
      {showM365 && m365Apps.length > 0 && (
        <>
          {showEntra && <SectionHeader text="Microsoft 365" />}
          {renderApps(m365Apps)}
        </>
      )}
      {showEntra && (
        <>
          {showM365 && <SectionHeader text="My applications" />}
          {entra.loading && !entra.data && <LoadingState rows={2} />}
          {entra.error && <ErrorState message={entra.error} onRetry={entra.reload} hint="My Apps requires the Directory.Read.All permission to be approved." />}
          {entra.data && entraApps.length === 0 && !entra.loading && (
            <EmptyState icon="AppIconDefault" title={q ? 'No matching apps' : 'No assigned applications'} message={q ? undefined : 'Applications assigned to you in Microsoft Entra will appear here.'} />
          )}
          {entra.data && entraApps.length > 0 && renderApps(entraApps)}
          {entra.data?.degraded && (
            <div style={{ fontSize: 11.5, color: 'var(--tpms-text-muted)', paddingTop: 6 }}>
              <Icon iconName="Info" /> App logos and direct launch links need Directory.Read.All; falling back to the My Apps portal.
            </div>
          )}
        </>
      )}
      <FooterLink href={myAppsUrl} text="Open My Apps portal" />
    </div>
  );
};

export const myAppsDefinition: IWidgetDefinition = {
  type: 'apps',
  title: 'My apps',
  description: 'Applications assigned to you in Microsoft Entra (My Apps) alongside Microsoft 365 launch tiles.',
  icon: 'AppIconDefaultList',
  category: 'productivity',
  defaultSize: 1,
  allowMultiple: false,
  permissions: ['Directory.Read.All'],
  settingsSchema: [
    { key: 'layout', label: 'Layout', type: 'choice', defaultValue: 'tiles', choices: [{ key: 'tiles', text: 'Tiles' }, { key: 'list', text: 'List' }] },
    { key: 'showSearch', label: 'Show search box', type: 'boolean', defaultValue: true },
    { key: 'showEntra', label: 'Show Entra assigned apps', type: 'boolean', defaultValue: true },
    { key: 'showM365', label: 'Show Microsoft 365 apps', type: 'boolean', defaultValue: true },
    {
      key: 'm365Apps',
      label: 'Microsoft 365 apps to show',
      type: 'text',
      defaultValue: 'outlook,teams,onedrive,sharepoint,word,excel,powerpoint,onenote,planner,todo,copilot',
      description: `Comma separated keys: ${M365_APPS.map(a => a.id).join(', ')}`
    }
  ],
  component: MyAppsWidget
};

export default MyAppsWidget;
