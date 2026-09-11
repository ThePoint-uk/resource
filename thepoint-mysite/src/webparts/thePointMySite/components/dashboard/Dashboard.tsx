import * as React from 'react';
import { DefaultButton, Dialog, DialogFooter, DialogType, Icon, MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize, TextField } from '@fluentui/react';
import type { IWidgetDefinition, IWidgetInstance } from '../../models';
import { useDashboard } from '../../context/DashboardContext';
import { getWidgetDefinition } from '../../widgets/registry';
import { DashboardHeader } from './DashboardHeader';
import { WidgetShell } from './WidgetShell';
import { WidgetCatalog } from './WidgetCatalog';
import { WidgetSettingsPanel } from './WidgetSettingsPanel';
import { WidgetErrorBoundary } from './WidgetErrorBoundary';
import styles from './Dashboard.module.scss';

interface IDragState {
  draggingId?: string;
  overId?: string;
  position?: 'before' | 'after';
}

function hexToRgba(hex: string, alpha: number): string {
  const clean: string = hex.replace('#', '');
  const full: string = clean.length === 3 ? clean.split('').map(c => c + c).join('') : clean;
  const num: number = parseInt(full, 16);
  if (isNaN(num) || full.length !== 6) {
    return `rgba(0, 120, 212, ${alpha})`;
  }
  return `rgba(${(num >> 16) & 255}, ${(num >> 8) & 255}, ${num & 255}, ${alpha})`;
}

export const Dashboard: React.FC = () => {
  const dash = useDashboard();
  const { props, state, ready, editMode, isDarkTheme, theme, loadError, loadSource, moveWidget, updateWidgetSettings, resetLayout, exportLayout, importLayout, refreshToken, graph, sp, context } = dash;
  const columns: number = Math.max(1, Math.min(4, props.columns || 3));
  const [catalogOpen, setCatalogOpen] = React.useState<boolean>(false);
  const [settingsId, setSettingsId] = React.useState<string | undefined>(undefined);
  const [resetOpen, setResetOpen] = React.useState<boolean>(false);
  const [ioOpen, setIoOpen] = React.useState<boolean>(false);
  const [ioText, setIoText] = React.useState<string>('');
  const [ioError, setIoError] = React.useState<string | undefined>(undefined);
  const [drag, setDrag] = React.useState<IDragState>({});

  const accent: string = props.accentColor || theme?.palette?.themePrimary || '#0078d4';
  const cssVars: Record<string, string | number | undefined> = {
    '--tpms-columns': columns,
    '--tpms-accent': accent,
    '--tpms-accent-soft': hexToRgba(accent, isDarkTheme ? 0.22 : 0.12),
    '--tpms-radius': `${props.cornerRadius ?? 8}px`
  };
  if (theme?.semanticColors) {
    cssVars['--tpms-text'] = theme.semanticColors.bodyText;
    cssVars['--tpms-text-muted'] = theme.semanticColors.bodySubtext;
    cssVars['--tpms-surface'] = theme.semanticColors.bodyBackground;
    cssVars['--tpms-surface-2'] = theme.palette?.neutralLighter;
    cssVars['--tpms-hover'] = theme.palette?.neutralLight;
    cssVars['--tpms-border'] = theme.semanticColors.bodyDivider;
  }

  const onDrop = (targetId: string): void => {
    if (!drag.draggingId || drag.draggingId === targetId) {
      setDrag({});
      return;
    }
    const targetIndex: number = state.widgets.findIndex(w => w.id === targetId);
    if (targetIndex >= 0) {
      moveWidget(drag.draggingId, drag.position === 'after' ? targetIndex + 1 : targetIndex);
    }
    setDrag({});
  };

  const openIo = (): void => {
    setIoText(exportLayout());
    setIoError(undefined);
    setIoOpen(true);
  };

  const applyImport = (): void => {
    if (importLayout(ioText)) {
      setIoOpen(false);
    } else {
      setIoError('That does not look like a valid layout. Expected JSON with a "widgets" array.');
    }
  };

  return (
    <div className={[styles.root, isDarkTheme ? styles.dark : '', props.density === 'compact' ? styles.compact : ''].join(' ')} style={cssVars as React.CSSProperties}>
      <DashboardHeader onAddWidget={() => setCatalogOpen(true)} onReset={() => setResetOpen(true)} onImportExport={openIo} />

      {loadError && (
        <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginBottom: 12 } }}>
          Your saved layout could not be loaded from {props.storageMode === 'onedrive' ? 'OneDrive' : 'the layouts list'} ({loadError}). Showing {loadSource === 'local' ? 'the copy saved in this browser' : 'the default layout'} instead.
        </MessageBar>
      )}

      {!ready ? (
        <div style={{ padding: 40, textAlign: 'center' }}>
          <Spinner size={SpinnerSize.large} label="Loading your dashboard…" />
        </div>
      ) : (
        <div className={styles.grid} onDragLeave={e => {
          if (e.currentTarget === e.target) {
            setDrag(d => ({ ...d, overId: undefined }));
          }
        }}>
          {state.widgets.length === 0 && (
            <div className={styles.emptyGrid}>
              <Icon iconName="ViewDashboard" styles={{ root: { fontSize: 32, color: 'var(--tpms-accent)' } }} />
              <div style={{ fontWeight: 600, marginTop: 8 }}>Your dashboard is empty</div>
              <div style={{ fontSize: 12.5, marginTop: 4 }}>Add widgets to build your personal workspace.</div>
              {dash.canCustomise && (
                <div style={{ marginTop: 12 }}>
                  <PrimaryButton
                    text="Add widget"
                    iconProps={{ iconName: 'Add' }}
                    onClick={() => {
                      dash.setEditMode(true);
                      setCatalogOpen(true);
                    }}
                  />
                </div>
              )}
            </div>
          )}
          {state.widgets.map((instance: IWidgetInstance, index: number) => {
            const definition: IWidgetDefinition | undefined = getWidgetDefinition(instance.type);
            if (!definition) {
              return undefined;
            }
            const Component = definition.component;
            return (
              <WidgetShell
                key={instance.id}
                instance={instance}
                definition={definition}
                index={index}
                total={state.widgets.length}
                columns={columns}
                onOpenSettings={setSettingsId}
                dragState={drag}
                onDragStart={id => setDrag({ draggingId: id })}
                onDragOver={(id, position) => setDrag(d => (d.overId === id && d.position === position ? d : { ...d, overId: id, position }))}
                onDrop={onDrop}
                onDragEnd={() => setDrag({})}
              >
                <WidgetErrorBoundary resetKey={`${instance.id}:${refreshToken}`}>
                  <Component
                    key={`${instance.id}:${refreshToken}`}
                    instance={instance}
                    settings={instance.settings}
                    context={context}
                    graph={graph}
                    sp={sp}
                    theme={theme}
                    isDarkTheme={isDarkTheme}
                    density={props.density || 'comfortable'}
                    editMode={editMode}
                    size={Math.min(instance.size, columns) as IWidgetInstance['size']}
                    updateSettings={patch => updateWidgetSettings(instance.id, patch)}
                  />
                </WidgetErrorBoundary>
              </WidgetShell>
            );
          })}
        </div>
      )}

      {ready && (
        <div className={styles.status}>
          <Icon iconName={props.storageMode === 'local' ? 'ThisPC' : 'CloudUpload'} />
          {props.storageMode === 'local' && 'Layout is saved in this browser.'}
          {props.storageMode === 'onedrive' && 'Layout is saved to your OneDrive and follows you across devices.'}
          {props.storageMode === 'list' && 'Layout is saved to SharePoint and follows you across devices.'}
          {props.lockLayout && ' Customisation has been locked by the page owner.'}
        </div>
      )}

      <WidgetCatalog isOpen={catalogOpen} onDismiss={() => setCatalogOpen(false)} />
      <WidgetSettingsPanel widgetId={settingsId} onDismiss={() => setSettingsId(undefined)} />

      <Dialog
        hidden={!resetOpen}
        onDismiss={() => setResetOpen(false)}
        dialogContentProps={{ type: DialogType.normal, title: 'Reset dashboard?', subText: 'This removes your customisations and restores the default layout set by the page owner.' }}
      >
        <DialogFooter>
          <PrimaryButton
            text="Reset"
            onClick={() => {
              resetLayout().catch(() => undefined);
              setResetOpen(false);
            }}
          />
          <DefaultButton text="Cancel" onClick={() => setResetOpen(false)} />
        </DialogFooter>
      </Dialog>

      <Dialog
        hidden={!ioOpen}
        onDismiss={() => setIoOpen(false)}
        minWidth={520}
        dialogContentProps={{ type: DialogType.normal, title: 'Import / export layout', subText: 'Copy this JSON to back up or share your layout, or paste a layout to import it.' }}
      >
        <TextField multiline rows={14} value={ioText} onChange={(_, v) => setIoText(v || '')} errorMessage={ioError} styles={{ field: { fontFamily: 'Consolas, monospace', fontSize: 12 } }} />
        <DialogFooter>
          <PrimaryButton text="Import" onClick={applyImport} />
          <DefaultButton
            text="Copy"
            onClick={() => {
              navigator.clipboard?.writeText(ioText).catch(() => undefined);
            }}
          />
          <DefaultButton text="Close" onClick={() => setIoOpen(false)} />
        </DialogFooter>
      </Dialog>
    </div>
  );
};
