import * as React from 'react';
import { DefaultButton, Dropdown, Panel, PanelType, PrimaryButton, SpinButton, TextField, Toggle } from '@fluentui/react';
import type { IWidgetDefinition, IWidgetInstance, IWidgetSettingField } from '../../models';
import { useDashboard } from '../../context/DashboardContext';
import { getWidgetDefinition } from '../../widgets/registry';

export interface IWidgetSettingsPanelProps {
  widgetId: string | undefined;
  onDismiss: () => void;
}

/** Generic settings editor driven by a widget definition's settings schema. */
export const WidgetSettingsPanel: React.FC<IWidgetSettingsPanelProps> = ({ widgetId, onDismiss }) => {
  const { state, updateWidget, updateWidgetSettings } = useDashboard();
  const instance: IWidgetInstance | undefined = state.widgets.find(w => w.id === widgetId);
  const definition: IWidgetDefinition | undefined = instance ? getWidgetDefinition(instance.type) : undefined;
  const [draft, setDraft] = React.useState<Record<string, unknown>>({});
  const [title, setTitle] = React.useState<string>('');

  React.useEffect(() => {
    if (instance) {
      setDraft({ ...instance.settings });
      setTitle(instance.title || '');
    }
  }, [instance?.id]);

  if (!instance || !definition) {
    return <Panel isOpen={false} onDismiss={onDismiss} />;
  }

  const fields: IWidgetSettingField[] = (definition.settingsSchema || []).filter(f => f.type !== 'links');

  const save = (): void => {
    updateWidgetSettings(instance.id, draft);
    const t: string = title.trim();
    updateWidget(instance.id, { title: t && t !== definition.title ? t : undefined });
    onDismiss();
  };

  const renderField = (field: IWidgetSettingField): React.ReactNode => {
    const value: unknown = draft[field.key] !== undefined ? draft[field.key] : field.defaultValue;
    switch (field.type) {
      case 'boolean':
        return <Toggle key={field.key} label={field.label} checked={value !== false && value !== undefined ? !!value : false} onText="On" offText="Off" onChange={(_, v) => setDraft(d => ({ ...d, [field.key]: !!v }))} />;
      case 'number':
        return (
          <SpinButton
            key={field.key}
            label={field.label}
            min={field.min}
            max={field.max}
            step={field.step || 1}
            value={String(value ?? '')}
            onChange={(_, v) => setDraft(d => ({ ...d, [field.key]: Number(v) }))}
            styles={{ root: { marginBottom: 8 } }}
          />
        );
      case 'choice':
        return (
          <Dropdown
            key={field.key}
            label={field.label}
            options={(field.choices || []).map(c => ({ key: c.key, text: c.text }))}
            selectedKey={value !== undefined ? String(value) : undefined}
            onChange={(_, o) => o && setDraft(d => ({ ...d, [field.key]: String(o.key) }))}
          />
        );
      case 'multiline':
        return <TextField key={field.key} label={field.label} multiline rows={5} value={String(value ?? '')} onChange={(_, v) => setDraft(d => ({ ...d, [field.key]: v || '' }))} description={field.description} />;
      default:
        return <TextField key={field.key} label={field.label} value={String(value ?? '')} onChange={(_, v) => setDraft(d => ({ ...d, [field.key]: v || '' }))} description={field.description} />;
    }
  };

  return (
    <Panel
      isOpen={!!widgetId}
      onDismiss={onDismiss}
      type={PanelType.smallFixedFar}
      headerText={`${definition.title} settings`}
      closeButtonAriaLabel="Close"
      isFooterAtBottom
      onRenderFooterContent={() => (
        <div style={{ display: 'flex', gap: 8 }}>
          <PrimaryButton text="Save" onClick={save} />
          <DefaultButton text="Cancel" onClick={onDismiss} />
        </div>
      )}
    >
      <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
        <TextField label="Title" value={title} placeholder={definition.title} onChange={(_, v) => setTitle(v || '')} description="Leave blank to use the default title." />
        {fields.map(renderField)}
        {fields.length === 0 && <div style={{ fontSize: 12.5, color: 'var(--tpms-text-muted)' }}>This widget has no additional settings.</div>}
      </div>
    </Panel>
  );
};
