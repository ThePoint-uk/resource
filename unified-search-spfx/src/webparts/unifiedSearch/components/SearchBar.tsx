import * as React from 'react';
import {
  SearchBox, Callout, DirectionalHint, Icon, Persona, PersonaSize, Text, ISearchBox
} from '@fluentui/react';
import { ISuggestions, IFileResult, IPersonResult } from '../models';
import { fileColor, fileIconName } from './styles';

export interface ISearchBarProps {
  initialQuery: string;
  enableSuggestions: boolean;
  onSearch: (query: string) => void;
  getSuggestions: (query: string) => Promise<ISuggestions>;
  onPickFile: (file: IFileResult) => void;
  onPickPerson: (person: IPersonResult) => void;
  recentSearches: string[];
}

export const SearchBar: React.FunctionComponent<ISearchBarProps> = props => {
  const [value, setValue] = React.useState<string>(props.initialQuery);
  const [suggestions, setSuggestions] = React.useState<ISuggestions | undefined>(undefined);
  const [showSuggestions, setShowSuggestions] = React.useState<boolean>(false);
  const containerRef = React.useRef<HTMLDivElement>(null);
  const searchBoxRef = React.useRef<ISearchBox>(null);
  const debounceRef = React.useRef<number | undefined>(undefined);
  const latestQueryRef = React.useRef<string>('');

  React.useEffect(() => { setValue(props.initialQuery); }, [props.initialQuery]);

  // "/" anywhere on the page focuses search, like Microsoft Search.
  React.useEffect(() => {
    const onKeyDown = (e: KeyboardEvent): void => {
      const target = e.target as HTMLElement;
      if (e.key === '/' && target.tagName !== 'INPUT' && target.tagName !== 'TEXTAREA') {
        e.preventDefault();
        if (searchBoxRef.current) { searchBoxRef.current.focus(); }
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, []);

  const queueSuggestions = (query: string): void => {
    latestQueryRef.current = query;
    if (debounceRef.current !== undefined) { window.clearTimeout(debounceRef.current); }
    if (!props.enableSuggestions || query.trim().length < 2) {
      setShowSuggestions(false);
      return;
    }
    debounceRef.current = window.setTimeout(() => {
      const requested = query;
      props.getSuggestions(requested.trim())
        .then(result => {
          if (latestQueryRef.current === requested &&
              (result.files.length > 0 || result.people.length > 0)) {
            setSuggestions(result);
            setShowSuggestions(true);
          } else if (latestQueryRef.current === requested) {
            setShowSuggestions(false);
          }
        })
        .catch(() => setShowSuggestions(false));
    }, 350);
  };

  const matchingRecent = props.recentSearches
    .filter(recent => value.trim().length >= 2 && recent.indexOf(value.trim().toLowerCase()) !== -1 && recent !== value.trim().toLowerCase())
    .slice(0, 3);

  const rowStyle: React.CSSProperties = {
    display: 'flex', alignItems: 'center', gap: 12, padding: '8px 16px', cursor: 'pointer'
  };

  return (
    <div ref={containerRef}>
      <SearchBox
        componentRef={searchBoxRef}
        placeholder="Search files, mail, events, people, sites and policies…"
        value={value}
        onChange={(_, newValue) => {
          const query = newValue || '';
          setValue(query);
          queueSuggestions(query);
        }}
        onSearch={query => { setShowSuggestions(false); props.onSearch((query || '').trim()); }}
        onClear={() => { setShowSuggestions(false); props.onSearch(''); }}
        styles={{
          root: {
            height: 52, fontSize: 16, borderRadius: 4, border: 'none',
            boxShadow: '0 1.6px 3.6px rgba(0,0,0,.132)'
          },
          field: { fontSize: 16 }
        }}
      />
      {showSuggestions && suggestions && containerRef.current &&
        <Callout
          target={containerRef.current}
          isBeakVisible={false}
          directionalHint={DirectionalHint.bottomLeftEdge}
          onDismiss={() => setShowSuggestions(false)}
          calloutWidth={containerRef.current.offsetWidth}
          preventDismissOnLostFocus={false}
          setInitialFocus={false}
          doNotLayer={false}
        >
          <div style={{ padding: '6px 0' }}>
            {matchingRecent.length > 0 && <>
              <Text variant="xSmall" block styles={{ root: { padding: '4px 16px', textTransform: 'uppercase', letterSpacing: '.5px' } }}>
                Recent searches
              </Text>
              {matchingRecent.map(recent =>
                <div key={recent} style={rowStyle} role="button" tabIndex={0}
                  onClick={() => { setShowSuggestions(false); setValue(recent); props.onSearch(recent); }}
                  onKeyDown={e => { if (e.key === 'Enter') { setShowSuggestions(false); props.onSearch(recent); } }}>
                  <Icon iconName="History" />
                  <span>{recent}</span>
                </div>)}
            </>}
            {suggestions.people.length > 0 && <>
              <Text variant="xSmall" block styles={{ root: { padding: '4px 16px', textTransform: 'uppercase', letterSpacing: '.5px' } }}>
                People
              </Text>
              {suggestions.people.map(person =>
                <div key={person.id} style={rowStyle} role="button" tabIndex={0}
                  onClick={() => { setShowSuggestions(false); props.onPickPerson(person); }}
                  onKeyDown={e => { if (e.key === 'Enter') { setShowSuggestions(false); props.onPickPerson(person); } }}>
                  <Persona text={person.displayName} size={PersonaSize.size24} hidePersonaDetails />
                  <span>{person.displayName}</span>
                  <Text variant="small" styles={{ root: { marginLeft: 'auto', opacity: .7 } }}>{person.jobTitle}</Text>
                </div>)}
            </>}
            {suggestions.files.length > 0 && <>
              <Text variant="xSmall" block styles={{ root: { padding: '4px 16px', textTransform: 'uppercase', letterSpacing: '.5px' } }}>
                Files
              </Text>
              {suggestions.files.map(file =>
                <div key={file.key} style={rowStyle} role="button" tabIndex={0}
                  onClick={() => { setShowSuggestions(false); props.onPickFile(file); }}
                  onKeyDown={e => { if (e.key === 'Enter') { setShowSuggestions(false); props.onPickFile(file); } }}>
                  <Icon iconName={fileIconName(file.fileType)} style={{ color: fileColor(file.fileType) }} />
                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{file.name}</span>
                  <Text variant="small" styles={{ root: { marginLeft: 'auto', opacity: .7, whiteSpace: 'nowrap' } }}>{file.siteName}</Text>
                </div>)}
            </>}
          </div>
        </Callout>}
    </div>
  );
};
