import * as React from 'react';
import { Icon, Panel, PanelType, SearchBox } from '@fluentui/react';
import type { IWidgetDefinition, WidgetCategory } from '../../models';
import { useDashboard } from '../../context/DashboardContext';
import { CATEGORY_LABELS, CATEGORY_ORDER } from '../../widgets/registry';
import styles from './Dashboard.module.scss';

export interface IWidgetCatalogProps {
  isOpen: boolean;
  onDismiss: () => void;
}

/** Side panel listing every available widget grouped by category. */
export const WidgetCatalog: React.FC<IWidgetCatalogProps> = ({ isOpen, onDismiss }) => {
  const { availableWidgets, state, addWidget, props, isPinned, unavailableTypes } = useDashboard();
  const [query, setQuery] = React.useState<string>('');
  const q: string = query.trim().toLowerCase();

  const groups: { category: WidgetCategory; items: IWidgetDefinition[] }[] = CATEGORY_ORDER.map(category => ({
    category,
    items: availableWidgets.filter(d => d.category === category && (!q || d.title.toLowerCase().indexOf(q) >= 0 || d.description.toLowerCase().indexOf(q) >= 0))
  })).filter(g => g.items.length > 0);

  return (
    <Panel isOpen={isOpen} onDismiss={onDismiss} type={PanelType.medium} headerText="Add a widget" closeButtonAriaLabel="Close">
      <div className={styles.catalogue}>
        <SearchBox placeholder="Search widgets" value={query} onChange={(_, v) => setQuery(v || '')} />
        {groups.length === 0 && <div className={styles.status}>No widgets match your search.</div>}
        {unavailableTypes.size > 0 && !q && (
          <div className={styles.status}>
            <Icon iconName="Info" /> {unavailableTypes.size} widget{unavailableTypes.size === 1 ? ' is' : 's are'} hidden because their permissions have not been approved yet.
          </div>
        )}
        {groups.map(group => (
          <div className={styles.catalogueSection} key={group.category}>
            <div className={styles.catalogueHeading}>{CATEGORY_LABELS[group.category]}</div>
            {group.items.map(def => {
              const alreadyAdded: boolean = def.allowMultiple === false && state.widgets.some(w => w.type === def.type);
              return (
                <button
                  type="button"
                  key={def.type}
                  className={styles.catalogueItem}
                  disabled={alreadyAdded}
                  onClick={() => {
                    addWidget(def.type);
                    onDismiss();
                  }}
                >
                  <span className={styles.catalogueIcon}>
                    <Icon iconName={def.icon} />
                  </span>
                  <span className={styles.catalogueBody}>
                    <span className={styles.catalogueTitle}>
                      {def.title}
                      {isPinned(def.type) && <span className={styles.requiredPill}><Icon iconName="Pinned" /> Required</span>}
                      {alreadyAdded && <span style={{ fontWeight: 400, color: 'var(--tpms-text-muted)' }}> · already on your dashboard</span>}
                    </span>
                    <span className={styles.catalogueDesc}>{def.description}</span>
                    {props.showPermissionHints && def.permissions && def.permissions.length > 0 && (
                      <span className={styles.cataloguePerms}>
                        <Icon iconName="Permissions" /> {def.permissions.join(', ')}
                      </span>
                    )}
                  </span>
                  <Icon iconName="Add" />
                </button>
              );
            })}
          </div>
        ))}
      </div>
    </Panel>
  );
};
