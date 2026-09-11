import * as React from 'react';
import type { IWidgetDefinition, IWidgetProps } from '../../models';

const COLOURS: Record<string, { bg: string; fg: string }> = {
  yellow: { bg: '#fff4ce', fg: '#323130' },
  green: { bg: '#dff6dd', fg: '#323130' },
  blue: { bg: '#deecf9', fg: '#323130' },
  pink: { bg: '#fde7e9', fg: '#323130' },
  neutral: { bg: 'var(--tpms-surface-2)', fg: 'var(--tpms-text)' }
};

/** Free-form personal notes saved with the layout. */
const NotesWidget: React.FC<IWidgetProps> = ({ settings, updateSettings, isDarkTheme }) => {
  const [text, setText] = React.useState<string>(String(settings.text || ''));
  const timer = React.useRef<number | undefined>(undefined);
  const colour = COLOURS[String(settings.colour || 'yellow')] || COLOURS.yellow;

  React.useEffect(() => {
    setText(String(settings.text || ''));
  }, [settings.text]);

  const onChange = (value: string): void => {
    setText(value);
    if (timer.current) {
      window.clearTimeout(timer.current);
    }
    timer.current = window.setTimeout(() => updateSettings({ text: value }), 600);
  };

  return (
    <textarea
      value={text}
      onChange={e => onChange(e.target.value)}
      placeholder="Jot down a note…"
      aria-label="Personal notes"
      style={{
        width: '100%',
        minHeight: 140,
        resize: 'vertical',
        border: 'none',
        borderRadius: 6,
        padding: 10,
        background: colour.bg,
        color: isDarkTheme && settings.colour !== 'neutral' ? '#323130' : colour.fg,
        font: 'inherit',
        fontSize: 13,
        lineHeight: 1.5,
        boxSizing: 'border-box'
      }}
    />
  );
};

export const notesDefinition: IWidgetDefinition = {
  type: 'notes',
  title: 'Notes',
  description: 'A personal sticky note that stays with your dashboard.',
  icon: 'QuickNote',
  category: 'personal',
  defaultSize: 1,
  allowMultiple: true,
  settingsSchema: [
    { key: 'colour', label: 'Colour', type: 'choice', defaultValue: 'yellow', choices: [{ key: 'yellow', text: 'Yellow' }, { key: 'green', text: 'Green' }, { key: 'blue', text: 'Blue' }, { key: 'pink', text: 'Pink' }, { key: 'neutral', text: 'Neutral' }] },
    { key: 'text', label: 'Text', type: 'multiline', defaultValue: '' }
  ],
  component: NotesWidget
};

export default NotesWidget;
