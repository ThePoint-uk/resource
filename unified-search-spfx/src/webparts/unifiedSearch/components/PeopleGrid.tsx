import * as React from 'react';
import {
  Persona, PersonaSize, PersonaPresence, Callout, DirectionalHint, IconButton,
  PrimaryButton, DefaultButton, Text, Icon, Spinner, SpinnerSize, Link
} from '@fluentui/react';
import { IPersonResult, IPersonOrg, PresenceAvailability } from '../models';
import { PeopleService } from '../services/PeopleService';
import { IUnifiedSearchClassNames } from './styles';

function toPersonaPresence(availability?: PresenceAvailability): PersonaPresence {
  switch (availability) {
    case 'Available': return PersonaPresence.online;
    case 'Busy': return PersonaPresence.busy;
    case 'DoNotDisturb': return PersonaPresence.dnd;
    case 'Away':
    case 'BeRightBack': return PersonaPresence.away;
    case 'Offline': return PersonaPresence.offline;
    default: return PersonaPresence.none;
  }
}

/** Teams deep links open the Teams client; the web part itself stays SharePoint-hosted. */
function teamsChatUrl(person: IPersonResult): string {
  return `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(person.mail || person.userPrincipalName || '')}`;
}

export interface IPersonCardProps {
  person: IPersonResult;
  target: HTMLElement;
  peopleService: PeopleService;
  onDismiss: () => void;
  onNavigate: (person: IPersonResult, target: HTMLElement) => void;
  onFilterDepartment?: (department: string) => void;
}

export const PersonCard: React.FunctionComponent<IPersonCardProps> = props => {
  const { person } = props;
  const [org, setOrg] = React.useState<IPersonOrg | undefined>(undefined);
  const [orgLoading, setOrgLoading] = React.useState<boolean>(true);

  React.useEffect(() => {
    let cancelled = false;
    setOrg(undefined);
    setOrgLoading(true);
    props.peopleService.getOrg(person.id)
      .then(result => { if (!cancelled) { setOrg(result); setOrgLoading(false); } })
      .catch(() => { if (!cancelled) { setOrgLoading(false); } });
    return () => { cancelled = true; };
  }, [person.id]);

  return (
    <Callout
      target={props.target}
      onDismiss={props.onDismiss}
      directionalHint={DirectionalHint.bottomLeftEdge}
      setInitialFocus
      styles={{ root: { width: 340, padding: 0 } }}
    >
      <div style={{ padding: '18px 20px 16px' }}>
        <Persona
          text={person.displayName}
          secondaryText={person.jobTitle}
          tertiaryText={[person.department, person.officeLocation].filter(Boolean).join(' · ')}
          size={PersonaSize.size72}
          presence={toPersonaPresence(person.presence)}
          imageUrl={person.photoUrl}
        />
        <div style={{ display: 'flex', gap: 8, margin: '14px 0' }}>
          {person.businessPhone &&
            <PrimaryButton
              iconProps={{ iconName: 'Phone' }}
              text="Call"
              href={`tel:${person.businessPhone.replace(/\s/g, '')}`}
              styles={{ root: { flex: 1 } }}
            />}
          <DefaultButton
            iconProps={{ iconName: 'TeamsLogo' }}
            text="Teams"
            href={teamsChatUrl(person)}
            target="_blank"
            rel="noopener"
            styles={{ root: { flex: 1 } }}
          />
          {person.mail &&
            <DefaultButton
              iconProps={{ iconName: 'Mail' }}
              text="Email"
              href={`mailto:${person.mail}`}
              styles={{ root: { flex: 1 } }}
            />}
        </div>
        <div style={{ borderTop: '1px solid rgba(0,0,0,.08)', paddingTop: 10, fontSize: 13 }}>
          {person.businessPhone &&
            <div style={{ padding: '3px 0' }}><Icon iconName="Phone" style={{ marginRight: 8 }} />{person.businessPhone}</div>}
          {person.mobilePhone &&
            <div style={{ padding: '3px 0' }}><Icon iconName="CellPhone" style={{ marginRight: 8 }} />{person.mobilePhone}</div>}
          {person.mail &&
            <div style={{ padding: '3px 0' }}>
              <Icon iconName="Mail" style={{ marginRight: 8 }} />
              <Link href={`mailto:${person.mail}`}>{person.mail}</Link>
            </div>}
          {person.officeLocation &&
            <div style={{ padding: '3px 0' }}><Icon iconName="MapPin" style={{ marginRight: 8 }} />{person.officeLocation}</div>}
          {person.department && props.onFilterDepartment &&
            <div style={{ padding: '3px 0' }}>
              <Icon iconName="Filter" style={{ marginRight: 8 }} />
              <Link onClick={() => props.onFilterDepartment && props.onFilterDepartment(person.department || '')}>
                Show everyone in {person.department}
              </Link>
            </div>}
        </div>
        <div style={{ borderTop: '1px solid rgba(0,0,0,.08)', marginTop: 10, paddingTop: 10 }}>
          <Text variant="xSmall" block styles={{ root: { textTransform: 'uppercase', letterSpacing: '.4px', marginBottom: 6 } }}>
            Organisation
          </Text>
          {orgLoading && <Spinner size={SpinnerSize.xSmall} />}
          {!orgLoading && org && org.manager &&
            <div
              role="button"
              tabIndex={0}
              style={{ padding: '3px 0', cursor: 'pointer' }}
              onClick={e => props.onNavigate(org.manager as IPersonResult, e.currentTarget as HTMLElement)}
              onKeyDown={e => { if (e.key === 'Enter') { props.onNavigate(org.manager as IPersonResult, e.currentTarget as HTMLElement); } }}
            >
              <Icon iconName="Up" style={{ marginRight: 8 }} />
              <strong>{org.manager.displayName}</strong>{org.manager.jobTitle ? ` · ${org.manager.jobTitle}` : ''}
            </div>}
          {!orgLoading && org && org.directReports.map(report =>
            <div
              key={report.id}
              role="button"
              tabIndex={0}
              style={{ padding: '3px 0', cursor: 'pointer' }}
              onClick={e => props.onNavigate(report, e.currentTarget as HTMLElement)}
              onKeyDown={e => { if (e.key === 'Enter') { props.onNavigate(report, e.currentTarget as HTMLElement); } }}
            >
              <Icon iconName="Down" style={{ marginRight: 8 }} />
              {report.displayName}{report.jobTitle ? ` · ${report.jobTitle}` : ''}
            </div>)}
          {!orgLoading && (!org || (!org.manager && org.directReports.length === 0)) &&
            <Text variant="small">No organisation information available.</Text>}
        </div>
      </div>
    </Callout>
  );
};

export interface IPeopleGridProps {
  people: IPersonResult[];
  classNames: IUnifiedSearchClassNames;
  peopleService: PeopleService;
  onFilterDepartment?: (department: string) => void;
}

export const PeopleGrid: React.FunctionComponent<IPeopleGridProps> = props => {
  const [selected, setSelected] = React.useState<{ person: IPersonResult; target: HTMLElement } | undefined>(undefined);

  return (
    <>
      <div className={props.classNames.peopleGrid}>
        {props.people.map(person =>
          <div
            key={person.id}
            className={props.classNames.personCard}
            role="button"
            tabIndex={0}
            onClick={e => setSelected({ person, target: e.currentTarget as HTMLElement })}
            onKeyDown={e => { if (e.key === 'Enter') { setSelected({ person, target: e.currentTarget as HTMLElement }); } }}
          >
            <Persona
              text={person.displayName}
              secondaryText={person.jobTitle}
              tertiaryText={[person.department, person.officeLocation].filter(Boolean).join(' · ')}
              size={PersonaSize.size48}
              presence={toPersonaPresence(person.presence)}
              imageUrl={person.photoUrl}
            />
            <div style={{ display: 'flex', gap: 2, marginTop: 10, borderTop: '1px solid rgba(0,0,0,.06)', paddingTop: 6 }}
              onClick={e => e.stopPropagation()}>
              {person.businessPhone &&
                <IconButton iconProps={{ iconName: 'Phone' }} title={`Call ${person.displayName}`}
                  href={`tel:${person.businessPhone.replace(/\s/g, '')}`} />}
              <IconButton iconProps={{ iconName: 'TeamsLogo' }} title="Chat in Teams"
                href={teamsChatUrl(person)} target="_blank" rel="noopener" />
              {person.mail &&
                <IconButton iconProps={{ iconName: 'Mail' }} title="Email" href={`mailto:${person.mail}`} />}
              <IconButton iconProps={{ iconName: 'ContactCard' }} title="Profile"
                onClick={e => setSelected({ person, target: (e.currentTarget as unknown) as HTMLElement })} />
            </div>
          </div>)}
      </div>
      {selected &&
        <PersonCard
          person={selected.person}
          target={selected.target}
          peopleService={props.peopleService}
          onDismiss={() => setSelected(undefined)}
          onNavigate={(person, target) => setSelected({ person, target })}
          onFilterDepartment={props.onFilterDepartment
            ? (department: string) => { setSelected(undefined); props.onFilterDepartment && props.onFilterDepartment(department); }
            : undefined}
        />}
    </>
  );
};
