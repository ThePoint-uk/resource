import * as React from 'react';
import { DefaultButton, IconButton, PrimaryButton, SearchBox, TooltipHost } from '@fluentui/react';
import { useDashboard } from '../../context/DashboardContext';
import { greetingForTime, formatDate } from '../../utils/format';
import styles from './Dashboard.module.scss';

export interface IDashboardHeaderProps {
  onAddWidget: () => void;
  onReset: () => void;
  onImportExport: () => void;
  onSaveAsDefault: () => void;
}

export const DashboardHeader: React.FC<IDashboardHeaderProps> = ({ onAddWidget, onReset, onImportExport, onSaveAsDefault }) => {
  const { props, context, editMode, setEditMode, canCustomise, saveAsDefault, refreshAll, sp } = useDashboard();
  const firstName: string = (context.pageContext.user.displayName || '').split(' ')[0];
  const [query, setQuery] = React.useState<string>('');

  const greeting: string | undefined =
    !props.showGreeting || props.greetingStyle === 'none'
      ? undefined
      : props.greetingStyle === 'welcome'
        ? `Welcome back, ${firstName}`
        : `${greetingForTime()}, ${firstName}`;

  const headerClass: string = props.headerStyle === 'accent' ? styles.headerAccent : props.headerStyle === 'gradient' ? styles.headerGradient : '';
  const inverted: boolean = props.headerStyle === 'accent' || props.headerStyle === 'gradient';

  const runSearch = (value: string | undefined): void => {
    const q: string = (value || '').trim();
    if (!q) {
      return;
    }
    window.open(`${sp.tenantUrl}/_layouts/15/search.aspx?q=${encodeURIComponent(q)}`, props.openLinksInNewTab ? '_blank' : '_self');
  };

  return (
    <div className={headerClass}>
      <div className={styles.header}>
        <div className={styles.headerText}>
          {props.showTitle && props.title && <div className={styles.subtitle} style={{ fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em', fontSize: 11 }}>{props.title}</div>}
          {greeting && <h2 className={styles.greeting}>{greeting}</h2>}
          {props.showDate && <div className={styles.subtitle}>{formatDate(new Date(), { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}</div>}
        </div>
        {props.showSearch && (
          <div className={styles.headerSearch}>
            <SearchBox placeholder="Search SharePoint" value={query} onChange={(_, v) => setQuery(v || '')} onSearch={runSearch} styles={inverted ? { root: { background: 'rgba(255,255,255,0.92)' } } : undefined} />
          </div>
        )}
        <div className={styles.headerActions}>
          <TooltipHost content="Refresh all widgets">
            <IconButton iconProps={{ iconName: 'Refresh' }} ariaLabel="Refresh all widgets" onClick={refreshAll} styles={inverted ? { root: { color: '#fff' }, rootHovered: { background: 'rgba(255,255,255,0.15)', color: '#fff' }, icon: { color: '#fff' } } : undefined} />
          </TooltipHost>
          {canCustomise && !editMode && (
            <DefaultButton text="Customise" iconProps={{ iconName: 'Edit' }} onClick={() => setEditMode(true)} styles={inverted ? { root: { background: 'rgba(255,255,255,0.92)', borderColor: 'transparent' } } : undefined} />
          )}
          {editMode && <PrimaryButton text="Done" iconProps={{ iconName: 'CheckMark' }} onClick={() => setEditMode(false)} />}
        </div>
      </div>
      {editMode && (
        <div className={styles.editBanner}>
          <span className={styles.bannerText}>
            <strong>Customising.</strong> Drag widgets by their handle, change widths with the 1 / 2 / 3 buttons, open settings from the cog, or add more from the catalogue. Changes save automatically.
          </span>
          <span className={styles.bannerActions}>
            <PrimaryButton text="Add widget" iconProps={{ iconName: 'Add' }} onClick={onAddWidget} />
            <DefaultButton text="Import / export" iconProps={{ iconName: 'Code' }} onClick={onImportExport} />
            <DefaultButton text="Reset" iconProps={{ iconName: 'Undo' }} onClick={onReset} />
            {saveAsDefault && (
              <TooltipHost content="Stores the current layout as the default for everyone, or for an audience you choose. Remember to republish the page.">
                <DefaultButton text="Set as page default" iconProps={{ iconName: 'Save' }} onClick={onSaveAsDefault} />
              </TooltipHost>
            )}
          </span>
        </div>
      )}
    </div>
  );
};
