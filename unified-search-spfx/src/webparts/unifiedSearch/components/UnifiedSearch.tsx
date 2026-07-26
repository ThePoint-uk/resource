import * as React from 'react';
import {
  Pivot, PivotItem, Dropdown, IDropdownOption, ActionButton, DefaultButton,
  MessageBar, MessageBarType, Spinner, SpinnerSize, Text, Icon, Link
} from '@fluentui/react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import {
  IUnifiedSearchConfig, VerticalKey, VERTICALS, isVerticalEnabled, SortMode,
  IFileResult, IMailResult, IEventResult, ISiteResult, IPersonResult,
  ISearchPage, IZeroState, ISuggestions, IPromotedResult
} from '../models';
import { GraphSearchService } from '../services/GraphSearchService';
import { PromotedResultsService } from '../services/PromotedResultsService';
import { PeopleService } from '../services/PeopleService';
import { FileService } from '../services/FileService';
import { getClassNames } from './styles';
import { SearchBar } from './SearchBar';
import { Refiners, IMailFilters, ICalendarFilters, IFacetBucket } from './Refiners';
import { FileRow, MailRow, EventRow, SiteCard } from './ResultRows';
import { PeopleGrid } from './PeopleGrid';
import { PreviewPanel } from './PreviewPanel';

export interface IUnifiedSearchProps {
  config: IUnifiedSearchConfig;
  theme: IReadonlyTheme | undefined;
  searchService: GraphSearchService;
  peopleService: PeopleService;
  fileService: FileService;
  promotedService: PromotedResultsService;
  userDisplayName: string;
  isEditMode: boolean;
}

const RECENT_KEY = 'unifiedSearch.recentSearches';

function readRecent(): string[] {
  try {
    const raw = window.localStorage.getItem(RECENT_KEY);
    return raw ? JSON.parse(raw) as string[] : [];
  } catch { return []; }
}

function writeRecent(query: string): string[] {
  const normalized = query.trim().toLowerCase();
  const next = [normalized, ...readRecent().filter(item => item !== normalized)].slice(0, 8);
  try { window.localStorage.setItem(RECENT_KEY, JSON.stringify(next)); } catch { /* private mode */ }
  return next;
}

interface IVerticalErrors { [key: string]: string | undefined; }

export const UnifiedSearch: React.FunctionComponent<IUnifiedSearchProps> = props => {
  const { config } = props;
  const classNames = React.useMemo(() => getClassNames(props.theme), [props.theme]);

  const [query, setQuery] = React.useState<string>('');
  const [vertical, setVertical] = React.useState<VerticalKey>('all');
  const [sort, setSort] = React.useState<SortMode>('relevance');

  const [fileTokens, setFileTokens] = React.useState<string[]>([]);
  const [mailFilters, setMailFilters] = React.useState<IMailFilters>({ dateKey: 'any', attachmentsOnly: false });
  const [calendarFilters, setCalendarFilters] = React.useState<ICalendarFilters>({ onlineOnly: false });
  const [selectedDepartments, setSelectedDepartments] = React.useState<string[]>([]);
  const [selectedOffices, setSelectedOffices] = React.useState<string[]>([]);

  const [filesPage, setFilesPage] = React.useState<ISearchPage<IFileResult> | undefined>(undefined);
  const [policiesPage, setPoliciesPage] = React.useState<ISearchPage<IFileResult> | undefined>(undefined);
  const [mailPage, setMailPage] = React.useState<ISearchPage<IMailResult> | undefined>(undefined);
  const [eventsPage, setEventsPage] = React.useState<ISearchPage<IEventResult> | undefined>(undefined);
  const [sitesPage, setSitesPage] = React.useState<ISearchPage<ISiteResult> | undefined>(undefined);
  const [people, setPeople] = React.useState<IPersonResult[] | undefined>(undefined);
  const [promoted, setPromoted] = React.useState<IPromotedResult[]>([]);

  const [loading, setLoading] = React.useState<boolean>(false);
  const [errors, setErrors] = React.useState<IVerticalErrors>({});
  const [spelling, setSpelling] = React.useState<string | undefined>(undefined);
  const [zeroState, setZeroState] = React.useState<IZeroState | undefined>(undefined);
  const [recentSearches, setRecentSearches] = React.useState<string[]>(readRecent());

  const [preview, setPreview] = React.useState<IFileResult | undefined>(undefined);
  const [toast, setToast] = React.useState<string | undefined>(undefined);

  const runToken = React.useRef<number>(0);
  const toastTimer = React.useRef<number | undefined>(undefined);

  const showToast = (text: string): void => {
    setToast(text);
    if (toastTimer.current !== undefined) { window.clearTimeout(toastTimer.current); }
    toastTimer.current = window.setTimeout(() => setToast(undefined), 3000);
  };

  /* ---------------- zero-query state ---------------- */

  React.useEffect(() => {
    if (query) { return; }
    let cancelled = false;
    Promise.all([
      props.fileService.getRecentFiles(5).catch((): IFileResult[] => []),
      props.fileService.getUpcomingEvents(4).catch((): IEventResult[] => [])
    ]).then(([recentFiles, upcomingEvents]) => {
      if (!cancelled) {
        setZeroState({ recentSearches: readRecent(), recentFiles, upcomingEvents });
      }
    }).catch(() => { /* zero state is best-effort */ });
    return () => { cancelled = true; };
  }, [query === '']);

  /* ---------------- main search ---------------- */

  const mailKql = (filters: IMailFilters): string | undefined => {
    const parts: string[] = [];
    if (filters.attachmentsOnly) { parts.push('hasattachments:true'); }
    if (filters.dateKey !== 'any') {
      const days = filters.dateKey === 'day' ? 1 : filters.dateKey === 'week' ? 7 : 30;
      const since = new Date(Date.now() - days * 86400000);
      parts.push(`received>=${since.toISOString().substring(0, 10)}`);
    }
    return parts.length > 0 ? parts.join(' ') : undefined;
  };

  const fetchFiles = (activeQuery: string, tokens: string[], sortMode: SortMode, token: number): void => {
    const scope = props.searchService.scopeKql(config.searchScope);
    props.searchService.searchFiles(activeQuery, {
      from: 0, size: config.pageSize, sort: sortMode,
      aggregationFilters: tokens, extraKql: scope,
      withRefiners: true, enableSpelling: config.enableSpelling
    }).then(page => {
      if (runToken.current === token) {
        setFilesPage(page);
        setSpelling(page.spellingSuggestion);
      }
    }).catch(() => {
      if (runToken.current === token) {
        setErrors(previous => ({ ...previous, files: 'File search failed. Check that Graph permissions are approved.' }));
      }
    });
  };

  const fetchPolicies = (activeQuery: string, tokens: string[], sortMode: SortMode, token: number): void => {
    props.searchService.searchFiles(activeQuery, {
      from: 0, size: config.pageSize, sort: sortMode,
      aggregationFilters: tokens, extraKql: config.policiesQuery,
      withRefiners: true, enableSpelling: false
    }).then(page => {
      if (runToken.current === token) { setPoliciesPage(page); }
    }).catch(() => {
      if (runToken.current === token) {
        setErrors(previous => ({ ...previous, policies: 'Policy search failed. Check the policies query template in the web part settings.' }));
      }
    });
  };

  const fetchMail = (activeQuery: string, filters: IMailFilters, token: number): void => {
    props.searchService.searchMail(activeQuery, 0, config.pageSize, mailKql(filters))
      .then(page => { if (runToken.current === token) { setMailPage(page); } })
      .catch(() => {
        if (runToken.current === token) {
          setErrors(previous => ({ ...previous, mail: 'Mail search failed. The Mail.Read permission may not be approved.' }));
        }
      });
  };

  const fetchEvents = (activeQuery: string, token: number): void => {
    props.searchService.searchEvents(activeQuery, 0, config.pageSize)
      .then(page => { if (runToken.current === token) { setEventsPage(page); } })
      .catch(() => {
        if (runToken.current === token) {
          setErrors(previous => ({ ...previous, calendar: 'Calendar search failed. The Calendars.Read permission may not be approved.' }));
        }
      });
  };

  const fetchSites = (activeQuery: string, token: number): void => {
    props.searchService.searchSites(activeQuery, 0, config.pageSize)
      .then(page => { if (runToken.current === token) { setSitesPage(page); } })
      .catch(() => {
        if (runToken.current === token) {
          setErrors(previous => ({ ...previous, sites: 'Site search failed.' }));
        }
      });
  };

  const fetchPeople = (activeQuery: string, token: number): void => {
    props.peopleService.search(activeQuery, 25)
      .then(results => {
        if (runToken.current !== token) { return; }
        setPeople(results);
        // Presence and photos are progressive enhancements; errors are fine.
        props.peopleService.getPresence(results).then(presenceMap => {
          if (runToken.current === token) {
            setPeople(current => (current || []).map(person => ({
              ...person,
              presence: presenceMap[person.id] || person.presence
            })));
          }
        }).catch(() => { /* Presence.Read.All not approved */ });
        results.slice(0, 12).forEach(person => {
          props.peopleService.getPhoto(person.id).then(photoUrl => {
            if (photoUrl && runToken.current === token) {
              setPeople(current => (current || []).map(item =>
                item.id === person.id ? { ...item, photoUrl } : item));
            }
          }).catch(() => { /* no photo */ });
        });
      })
      .catch(() => {
        if (runToken.current === token) {
          setErrors(previous => ({ ...previous, people: 'People search failed. The User.Read.All permission may not be approved.' }));
        }
      });
  };

  const fetchPromoted = (activeQuery: string, token: number): void => {
    props.promotedService.getMatches(activeQuery, config.promotedListTitle, config.promotedWebUrl)
      .then(matches => { if (runToken.current === token) { setPromoted(matches); } })
      .catch(() => { /* promoted results are optional */ });
  };

  const runSearch = (activeQuery: string): void => {
    runToken.current += 1;
    const token = runToken.current;
    setErrors({});
    setSpelling(undefined);
    setFilesPage(undefined); setPoliciesPage(undefined); setMailPage(undefined);
    setEventsPage(undefined); setSitesPage(undefined); setPeople(undefined);
    setPromoted([]);
    if (!activeQuery) { setLoading(false); return; }

    setLoading(true);
    const work: Promise<void>[] = [];
    const track = (fn: () => void): void => { work.push(Promise.resolve().then(fn)); };
    if (config.enablePromoted) { track(() => fetchPromoted(activeQuery, token)); }
    if (config.showFiles) { track(() => fetchFiles(activeQuery, [], sort, token)); }
    if (config.showPolicies) { track(() => fetchPolicies(activeQuery, [], sort, token)); }
    if (config.showMail) { track(() => fetchMail(activeQuery, mailFilters, token)); }
    if (config.showCalendar) { track(() => fetchEvents(activeQuery, token)); }
    if (config.showSites) { track(() => fetchSites(activeQuery, token)); }
    if (config.showPeople) { track(() => fetchPeople(activeQuery, token)); }
    // The individual fetches resolve independently; stop the global spinner once kicked off.
    Promise.all(work).then(() => {
      window.setTimeout(() => { if (runToken.current === token) { setLoading(false); } }, 400);
    }).catch(() => setLoading(false));
  };

  const onSearch = (newQuery: string): void => {
    setQuery(newQuery);
    setFileTokens([]);
    setSelectedDepartments([]);
    setSelectedOffices([]);
    if (newQuery) { setRecentSearches(writeRecent(newQuery)); }
    runSearch(newQuery);
  };

  // Re-run the file/policy legs when refiner tokens or sort change.
  React.useEffect(() => {
    if (!query) { return; }
    runToken.current += 1;
    const token = runToken.current;
    if (config.showFiles) { setFilesPage(undefined); fetchFiles(query, fileTokens, sort, token); }
    if (config.showPolicies) { setPoliciesPage(undefined); fetchPolicies(query, fileTokens, sort, token); }
  }, [fileTokens.join('|'), sort]);

  // Re-run mail when its filters change.
  React.useEffect(() => {
    if (!query || !config.showMail) { return; }
    runToken.current += 1;
    const token = runToken.current;
    setMailPage(undefined);
    fetchMail(query, mailFilters, token);
  }, [mailFilters.dateKey, mailFilters.attachmentsOnly]);

  /* ---------------- derived data ---------------- */

  const filteredPeople = React.useMemo(() => {
    let result = people || [];
    if (selectedDepartments.length > 0) {
      result = result.filter(person => person.department && selectedDepartments.indexOf(person.department) !== -1);
    }
    if (selectedOffices.length > 0) {
      result = result.filter(person => person.officeLocation && selectedOffices.indexOf(person.officeLocation) !== -1);
    }
    return result;
  }, [people, selectedDepartments, selectedOffices]);

  const facet = (items: IPersonResult[], selector: (person: IPersonResult) => string | undefined): IFacetBucket[] => {
    const counts: { [value: string]: number } = {};
    items.forEach(person => {
      const value = selector(person);
      if (value) { counts[value] = (counts[value] || 0) + 1; }
    });
    return Object.keys(counts).sort().map(value => ({ value, count: counts[value] }));
  };

  const departmentFacets = React.useMemo(() => facet(people || [], person => person.department), [people]);
  const officeFacets = React.useMemo(() => facet(people || [], person => person.officeLocation), [people]);

  const filteredEvents = React.useMemo(() => {
    const items = eventsPage ? eventsPage.items : [];
    return calendarFilters.onlineOnly ? items.filter(event => event.isOnline) : items;
  }, [eventsPage, calendarFilters.onlineOnly]);

  const sortedMail = React.useMemo(() => {
    const items = mailPage ? mailPage.items.slice() : [];
    if (sort === 'newest') { items.sort((a, b) => (b.received || '').localeCompare(a.received || '')); }
    if (sort === 'oldest') { items.sort((a, b) => (a.received || '').localeCompare(b.received || '')); }
    return items;
  }, [mailPage, sort]);

  const totalResults =
    (filesPage ? filesPage.total : 0) + (mailPage ? mailPage.total : 0) +
    (eventsPage ? eventsPage.total : 0) + (people ? people.length : 0) +
    (sitesPage ? sitesPage.total : 0) + (policiesPage ? policiesPage.total : 0);

  const countFor = (key: VerticalKey): number | undefined => {
    switch (key) {
      case 'all': return totalResults;
      case 'files': return filesPage ? filesPage.total : undefined;
      case 'mail': return mailPage ? mailPage.total : undefined;
      case 'calendar': return eventsPage ? eventsPage.total : undefined;
      case 'people': return people ? people.length : undefined;
      case 'sites': return sitesPage ? sitesPage.total : undefined;
      case 'policies': return policiesPage ? policiesPage.total : undefined;
      default: return undefined;
    }
  };

  /* ---------------- actions ---------------- */

  const copyLink = (file: IFileResult): void => {
    props.fileService.createViewLink(file)
      .then(url => navigator.clipboard.writeText(url))
      .then(() => showToast('Link copied. People in your organisation with the link can view.'))
      .catch(() => {
        navigator.clipboard.writeText(file.webUrl)
          .then(() => showToast('Direct link copied.'))
          .catch(() => showToast('Could not copy the link.'));
      });
  };

  const shareFromRow = (file: IFileResult): void => { setPreview(file); };

  /* ---------------- render helpers ---------------- */

  const sortOptions: IDropdownOption[] = [
    { key: 'relevance', text: 'Relevance' },
    { key: 'newest', text: 'Newest' },
    { key: 'oldest', text: 'Oldest' }
  ];

  const sectionHeader = (icon: string, title: string, seeAllKey?: VerticalKey): React.ReactNode => (
    <div className={classNames.sectionHeader}>
      <Icon iconName={icon} />
      <span>{title}</span>
      {seeAllKey &&
        <ActionButton
          styles={{ root: { marginLeft: 'auto', height: 24 } }}
          onClick={() => setVertical(seeAllKey)}
        >
          See all
        </ActionButton>}
    </div>
  );

  const errorBar = (key: string): React.ReactNode =>
    errors[key]
      ? <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginBottom: 10 } }}>{errors[key]}</MessageBar>
      : null;

  const fileList = (files: IFileResult[]): React.ReactNode => (
    <div className={classNames.card}>
      {files.map(file =>
        <FileRow
          key={file.key}
          file={file}
          classNames={classNames}
          openInNewTab={config.openInNewTab}
          enablePreview={config.enablePreviews}
          onPreview={setPreview}
          onShare={shareFromRow}
          onCopyLink={copyLink}
        />)}
    </div>
  );

  const emptyVertical = (label: string): React.ReactNode => (
    <div className={classNames.emptyState}>
      <Icon iconName="SearchIssue" style={{ fontSize: 32 }} />
      <Text variant="large" block styles={{ root: { margin: '10px 0 4px', fontWeight: 600 } }}>
        No {label} match "{query}"
      </Text>
      <Text variant="medium">Try different keywords or remove a filter.</Text>
    </div>
  );

  /* ---------------- vertical bodies ---------------- */

  const renderAll = (): React.ReactNode => (
    <>
      {promoted.map(promotedResult =>
        <div className={classNames.answerCard} key={promotedResult.title}>
          <Text variant="xSmall" styles={{ root: {
            fontWeight: 700, letterSpacing: '.5px', textTransform: 'uppercase', opacity: .75
          } }}>
            <Icon iconName="SingleBookmark" style={{ fontSize: 11, marginRight: 5 }} />
            Promoted result
          </Text>
          <Text variant="large" block styles={{ root: { fontWeight: 600, margin: '6px 0 4px' } }}>
            {promotedResult.title}
          </Text>
          {promotedResult.description &&
            <Text variant="medium" block styles={{ root: { marginBottom: 10, whiteSpace: 'pre-wrap' } }}>
              {promotedResult.description}
            </Text>}
          {promotedResult.linkUrl &&
            <DefaultButton
              primary
              text={promotedResult.buttonText || 'Open'}
              iconProps={{ iconName: 'OpenInNewWindow' }}
              onClick={() => window.open(
                promotedResult.linkUrl, config.openInNewTab ? '_blank' : '_self', 'noopener')}
            />}
        </div>)}
      {config.showPeople && filteredPeople.length > 0 &&
        <div className={classNames.section}>
          {sectionHeader('People', 'People', 'people')}
          <PeopleGrid
            people={filteredPeople.slice(0, 4)}
            classNames={classNames}
            peopleService={props.peopleService}
            onFilterDepartment={department => { setSelectedDepartments([department]); setVertical('people'); }}
          />
        </div>}
      {config.showFiles && <>
        {errorBar('files')}
        {filesPage && filesPage.items.length > 0 &&
          <div className={classNames.section}>
            {sectionHeader('Documentation', 'Files', 'files')}
            {fileList(filesPage.items.slice(0, 4))}
          </div>}
      </>}
      {config.showPolicies && policiesPage && policiesPage.items.length > 0 &&
        <div className={classNames.section}>
          {sectionHeader('Shield', 'Policies', 'policies')}
          {fileList(policiesPage.items.slice(0, 2))}
        </div>}
      {config.showMail && <>
        {errorBar('mail')}
        {sortedMail.length > 0 &&
          <div className={classNames.section}>
            {sectionHeader('Mail', 'Mail', 'mail')}
            <div className={classNames.card}>
              {sortedMail.slice(0, 3).map(mail =>
                <MailRow key={mail.key} mail={mail} classNames={classNames} openInNewTab={config.openInNewTab} />)}
            </div>
          </div>}
      </>}
      {config.showCalendar && filteredEvents.length > 0 &&
        <div className={classNames.section}>
          {sectionHeader('Calendar', 'Calendar', 'calendar')}
          <div className={classNames.card}>
            {filteredEvents.slice(0, 2).map(event =>
              <EventRow key={event.key} event={event} classNames={classNames} openInNewTab={config.openInNewTab} />)}
          </div>
        </div>}
      {config.showSites && sitesPage && sitesPage.items.length > 0 &&
        <div className={classNames.section}>
          {sectionHeader('Globe', 'Sites', 'sites')}
          <div className={classNames.siteGrid}>
            {sitesPage.items.slice(0, 2).map(site =>
              <SiteCard key={site.key} site={site} classNames={classNames} openInNewTab={config.openInNewTab} />)}
          </div>
        </div>}
      {!loading && totalResults === 0 && emptyVertical('results')}
    </>
  );

  const renderVertical = (): React.ReactNode => {
    switch (vertical) {
      case 'files':
        if (errors.files) { return errorBar('files'); }
        if (!filesPage) { return <Spinner size={SpinnerSize.large} styles={{ root: { padding: 40 } }} />; }
        if (filesPage.items.length === 0) { return emptyVertical('files'); }
        return fileList(filesPage.items);
      case 'policies':
        if (errors.policies) { return errorBar('policies'); }
        if (!policiesPage) { return <Spinner size={SpinnerSize.large} styles={{ root: { padding: 40 } }} />; }
        if (policiesPage.items.length === 0) { return emptyVertical('policies'); }
        return fileList(policiesPage.items);
      case 'mail':
        if (errors.mail) { return errorBar('mail'); }
        if (!mailPage) { return <Spinner size={SpinnerSize.large} styles={{ root: { padding: 40 } }} />; }
        if (sortedMail.length === 0) { return emptyVertical('messages'); }
        return (
          <div className={classNames.card}>
            {sortedMail.map(mail =>
              <MailRow key={mail.key} mail={mail} classNames={classNames} openInNewTab={config.openInNewTab} />)}
          </div>);
      case 'calendar':
        if (errors.calendar) { return errorBar('calendar'); }
        if (!eventsPage) { return <Spinner size={SpinnerSize.large} styles={{ root: { padding: 40 } }} />; }
        if (filteredEvents.length === 0) { return emptyVertical('events'); }
        return (
          <div className={classNames.card}>
            {filteredEvents.map(event =>
              <EventRow key={event.key} event={event} classNames={classNames} openInNewTab={config.openInNewTab} />)}
          </div>);
      case 'people':
        if (errors.people) { return errorBar('people'); }
        if (!people) { return <Spinner size={SpinnerSize.large} styles={{ root: { padding: 40 } }} />; }
        if (filteredPeople.length === 0) { return emptyVertical('people'); }
        return (
          <>
            <div className={classNames.chipRow}>
              <DefaultButton
                text="All departments"
                primary={selectedDepartments.length === 0}
                onClick={() => setSelectedDepartments([])}
                styles={{ root: { borderRadius: 14, height: 28, padding: '0 12px' } }}
              />
              {departmentFacets.map(bucket =>
                <DefaultButton
                  key={bucket.value}
                  text={`${bucket.value} (${bucket.count})`}
                  primary={selectedDepartments.indexOf(bucket.value) !== -1}
                  onClick={() => setSelectedDepartments(
                    selectedDepartments.indexOf(bucket.value) !== -1
                      ? selectedDepartments.filter(value => value !== bucket.value)
                      : [...selectedDepartments, bucket.value])}
                  styles={{ root: { borderRadius: 14, height: 28, padding: '0 12px' } }}
                />)}
            </div>
            <PeopleGrid
              people={filteredPeople}
              classNames={classNames}
              peopleService={props.peopleService}
              onFilterDepartment={department => setSelectedDepartments([department])}
            />
          </>);
      case 'sites':
        if (errors.sites) { return errorBar('sites'); }
        if (!sitesPage) { return <Spinner size={SpinnerSize.large} styles={{ root: { padding: 40 } }} />; }
        if (sitesPage.items.length === 0) { return emptyVertical('sites'); }
        return (
          <div className={classNames.siteGrid}>
            {sitesPage.items.map(site =>
              <SiteCard key={site.key} site={site} classNames={classNames} openInNewTab={config.openInNewTab} />)}
          </div>);
      default:
        return renderAll();
    }
  };

  const renderZeroState = (): React.ReactNode => (
    <div style={{ maxWidth: 1320, margin: '0 auto', padding: '18px 24px 48px' }}>
      {zeroState && zeroState.recentSearches.length > 0 &&
        <div className={classNames.section}>
          {sectionHeader('History', 'Your recent searches')}
          <div className={classNames.chipRow}>
            {zeroState.recentSearches.map(recent =>
              <DefaultButton
                key={recent}
                text={recent}
                iconProps={{ iconName: 'Search' }}
                onClick={() => onSearch(recent)}
                styles={{ root: { borderRadius: 14, height: 30 } }}
              />)}
          </div>
        </div>}
      {zeroState && zeroState.recentFiles.length > 0 &&
        <div className={classNames.section}>
          {sectionHeader('Recent', 'Pick up where you left off')}
          <div className={classNames.card}>
            {zeroState.recentFiles.map(file =>
              <FileRow
                key={file.key}
                file={file}
                classNames={classNames}
                openInNewTab={config.openInNewTab}
                enablePreview={false}
                onPreview={setPreview}
                onShare={shareFromRow}
                onCopyLink={copyLink}
              />)}
          </div>
        </div>}
      {zeroState && zeroState.upcomingEvents.length > 0 &&
        <div className={classNames.section}>
          {sectionHeader('Calendar', 'Coming up')}
          <div className={classNames.card}>
            {zeroState.upcomingEvents.map(event =>
              <EventRow key={event.key} event={event} classNames={classNames} openInNewTab={config.openInNewTab} />)}
          </div>
        </div>}
      {!zeroState &&
        <Spinner size={SpinnerSize.large} styles={{ root: { padding: 40 } }} />}
    </div>
  );

  /* ---------------- shell ---------------- */

  const showRefiners = config.filterPanelPosition !== 'hidden' && vertical !== 'sites';
  const refinersForVertical = vertical === 'policies'
    ? (policiesPage ? policiesPage.refiners : undefined)
    : (filesPage ? filesPage.refiners : undefined);

  return (
    <div className={classNames.root}>
      <div className={classNames.hero}>
        <div className={classNames.heroInner}>
          <div className={classNames.searchRow}>
            <SearchBar
              initialQuery={query}
              enableSuggestions={config.enableSuggestions}
              onSearch={onSearch}
              getSuggestions={(text): Promise<ISuggestions> =>
                props.searchService.suggest(text, (q, top) => props.peopleService.search(q, top))}
              onPickFile={file => {
                if (config.enablePreviews && file.driveId) { setPreview(file); }
                else { window.open(file.webUrl, config.openInNewTab ? '_blank' : '_self', 'noopener'); }
              }}
              onPickPerson={person => { onSearch(person.displayName); setVertical('people'); }}
              recentSearches={recentSearches}
            />
          </div>
        </div>
        <div className={classNames.pivotBar}>
          <Pivot
            selectedKey={vertical}
            onLinkClick={item => { if (item && item.props.itemKey) { setVertical(item.props.itemKey as VerticalKey); } }}
            overflowBehavior="menu"
          >
            {VERTICALS
              .filter(definition => isVerticalEnabled(config, definition.key))
              .map(definition => {
                const count = query ? countFor(definition.key) : undefined;
                return (
                  <PivotItem
                    key={definition.key}
                    itemKey={definition.key}
                    headerText={definition.label}
                    itemIcon={definition.icon}
                    itemCount={count}
                  />
                );
              })}
          </Pivot>
        </div>
      </div>

      {!query && renderZeroState()}

      {query && <>
        <div style={{ maxWidth: 1320, margin: '0 auto', padding: '0 24px' }}>
          <div className={classNames.statusLine}>
            {loading
              ? <Spinner size={SpinnerSize.small} labelPosition="right" label="Searching…" />
              : <span><strong>{countFor(vertical) !== undefined ? countFor(vertical) : '…'}</strong> results for "{query}"</span>}
            {spelling && config.enableSpelling &&
              <span>
                Did you mean <Link onClick={() => onSearch(spelling)}><strong>{spelling}</strong></Link>?
              </span>}
            <div style={{ flex: 1 }} />
            {(vertical === 'files' || vertical === 'policies' || vertical === 'mail' || vertical === 'all') &&
              <Dropdown
                options={sortOptions}
                selectedKey={sort}
                onChange={(_, option) => option && setSort(option.key as SortMode)}
                styles={{ dropdown: { width: 120 } }}
                ariaLabel="Sort results"
              />}
          </div>
        </div>
        <div className={
          classNames.layout + (config.filterPanelPosition === 'right' ? ' ' + classNames.layoutRight : '')
        }>
          {showRefiners &&
            <div className={classNames.refiners}>
              <Refiners
                classNames={classNames}
                vertical={vertical}
                fileRefiners={refinersForVertical}
                selectedFileTokens={fileTokens}
                onToggleFileToken={token => setFileTokens(
                  fileTokens.indexOf(token) !== -1
                    ? fileTokens.filter(existing => existing !== token)
                    : [...fileTokens, token])}
                onClearFileTokens={() => setFileTokens([])}
                mailFilters={mailFilters}
                onMailFiltersChange={setMailFilters}
                calendarFilters={calendarFilters}
                onCalendarFiltersChange={setCalendarFilters}
                departmentFacets={departmentFacets}
                officeFacets={officeFacets}
                selectedDepartments={selectedDepartments}
                selectedOffices={selectedOffices}
                onToggleDepartment={department => setSelectedDepartments(
                  selectedDepartments.indexOf(department) !== -1
                    ? selectedDepartments.filter(value => value !== department)
                    : [...selectedDepartments, department])}
                onToggleOffice={office => setSelectedOffices(
                  selectedOffices.indexOf(office) !== -1
                    ? selectedOffices.filter(value => value !== office)
                    : [...selectedOffices, office])}
              />
            </div>}
          <div className={classNames.results}>
            {renderVertical()}
          </div>
        </div>
      </>}

      {preview &&
        <PreviewPanel
          file={preview}
          files={(vertical === 'policies' ? (policiesPage && policiesPage.items) : (filesPage && filesPage.items)) || [preview]}
          fileService={props.fileService}
          openInNewTab={config.openInNewTab}
          onDismiss={() => setPreview(undefined)}
          onNavigate={setPreview}
          onCopyLink={copyLink}
        />}

      {toast &&
        <div style={{
          position: 'fixed', bottom: 24, left: '50%', transform: 'translateX(-50%)',
          background: '#242424', color: '#fff', borderRadius: 6, padding: '10px 18px',
          fontSize: 13, zIndex: 1000000, boxShadow: '0 6.4px 14.4px rgba(0,0,0,.3)'
        }} role="status">
          {toast}
        </div>}
    </div>
  );
};
