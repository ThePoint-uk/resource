import * as React from 'react';
import { Icon, IconButton, TooltipHost } from '@fluentui/react';
import type { IWidgetDefinition, IWidgetInstance, WidgetSize } from '../../models';
import { useDashboard } from '../../context/DashboardContext';
import styles from './Dashboard.module.scss';

export interface IWidgetShellProps {
  instance: IWidgetInstance;
  definition: IWidgetDefinition;
  index: number;
  total: number;
  columns: number;
  onOpenSettings: (id: string) => void;
  dragState: { draggingId?: string; overId?: string; position?: 'before' | 'after' };
  onDragStart: (id: string) => void;
  onDragOver: (id: string, position: 'before' | 'after') => void;
  onDrop: (id: string) => void;
  onDragEnd: () => void;
  children?: React.ReactNode;
}

/** Card chrome around a widget: title, actions, drag handle and edit-mode controls. */
export const WidgetShell: React.FC<IWidgetShellProps> = (props) => {
  const { instance, definition, index, total, columns, onOpenSettings, dragState, onDragStart, onDragOver, onDrop, onDragEnd, children } = props;
  const { editMode, removeWidget, updateWidget, moveWidget, props: wpProps } = useDashboard();
  const [renaming, setRenaming] = React.useState<boolean>(false);
  const [titleDraft, setTitleDraft] = React.useState<string>(instance.title || definition.title);
  const ref = React.useRef<HTMLDivElement>(null);

  const cardStyleClass: string = wpProps.cardStyle === 'outline' ? styles.styleOutline : wpProps.cardStyle === 'flat' ? styles.styleFlat : styles.styleShadow;
  const isDragging: boolean = dragState.draggingId === instance.id;
  const dropClass: string = dragState.overId === instance.id && dragState.draggingId !== instance.id ? (dragState.position === 'before' ? styles.dropBefore : styles.dropAfter) : '';
  const allowedSizes: WidgetSize[] = (definition.sizes || [1, 2, 3, 4]).filter(s => s <= columns) as WidgetSize[];
  const hasSettings: boolean = true;

  const commitRename = (): void => {
    const next: string = titleDraft.trim();
    updateWidget(instance.id, { title: next && next !== definition.title ? next : undefined });
    setRenaming(false);
  };

  return (
    <section
      ref={ref}
      className={[styles.widget, cardStyleClass, editMode ? styles.widgetEditing : '', instance.collapsed ? styles.widgetCollapsed : '', isDragging ? styles.dragging : '', dropClass].join(' ')}
      style={{ '--tpms-span': Math.min(instance.size, columns) } as React.CSSProperties}
      aria-label={instance.title || definition.title}
      onDragOver={e => {
        if (!editMode || !dragState.draggingId || dragState.draggingId === instance.id) {
          return;
        }
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        const rect: DOMRect = (e.currentTarget as HTMLElement).getBoundingClientRect();
        const horizontal: boolean = rect.width < rect.height * 2.5 || columns > 1;
        const position: 'before' | 'after' = horizontal ? (e.clientX < rect.left + rect.width / 2 ? 'before' : 'after') : e.clientY < rect.top + rect.height / 2 ? 'before' : 'after';
        onDragOver(instance.id, position);
      }}
      onDrop={e => {
        if (!editMode || !dragState.draggingId) {
          return;
        }
        e.preventDefault();
        onDrop(instance.id);
      }}
    >
      <header className={styles.widgetHeader}>
        {editMode && (
          <span
            className={styles.dragHandle}
            draggable
            title="Drag to move"
            aria-label="Drag to move"
            role="button"
            tabIndex={0}
            onDragStart={e => {
              e.dataTransfer.effectAllowed = 'move';
              e.dataTransfer.setData('text/plain', instance.id);
              if (ref.current) {
                e.dataTransfer.setDragImage(ref.current, 20, 20);
              }
              onDragStart(instance.id);
            }}
            onDragEnd={onDragEnd}
            onKeyDown={e => {
              if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
                e.preventDefault();
                moveWidget(instance.id, Math.max(0, index - 1));
              } else if (e.key === 'ArrowRight' || e.key === 'ArrowDown') {
                e.preventDefault();
                moveWidget(instance.id, Math.min(total, index + 2));
              }
            }}
          >
            <Icon iconName="GripperDotsVertical" />
          </span>
        )}
        <Icon iconName={definition.icon} className={styles.widgetIcon} aria-hidden="true" />
        {renaming ? (
          <input
            className={styles.widgetTitleInput}
            value={titleDraft}
            autoFocus
            aria-label="Widget title"
            onChange={e => setTitleDraft(e.target.value)}
            onBlur={commitRename}
            onKeyDown={e => {
              if (e.key === 'Enter') {
                commitRename();
              } else if (e.key === 'Escape') {
                setTitleDraft(instance.title || definition.title);
                setRenaming(false);
              }
            }}
          />
        ) : (
          <h3 className={styles.widgetTitle} onDoubleClick={() => editMode && setRenaming(true)} title={instance.title || definition.title}>
            {instance.title || definition.title}
          </h3>
        )}
        <div className={styles.widgetActions}>
          {editMode && allowedSizes.length > 1 && (
            <span className={styles.sizeButtons} role="group" aria-label="Widget width">
              {allowedSizes.map(s => (
                <TooltipHost content={`${s} column${s > 1 ? 's' : ''} wide`} key={s}>
                  <button type="button" className={instance.size === s ? styles.sizeActive : ''} onClick={() => updateWidget(instance.id, { size: s })} aria-pressed={instance.size === s}>
                    {s}
                  </button>
                </TooltipHost>
              ))}
            </span>
          )}
          {editMode && (
            <>
              <IconButton iconProps={{ iconName: 'ChevronLeft' }} title="Move earlier" ariaLabel="Move earlier" disabled={index === 0} onClick={() => moveWidget(instance.id, Math.max(0, index - 1))} styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 11 } }} />
              <IconButton iconProps={{ iconName: 'ChevronRight' }} title="Move later" ariaLabel="Move later" disabled={index >= total - 1} onClick={() => moveWidget(instance.id, Math.min(total, index + 2))} styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 11 } }} />
              <IconButton iconProps={{ iconName: 'Rename' }} title="Rename" ariaLabel="Rename widget" onClick={() => setRenaming(true)} styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 13 } }} />
              {hasSettings && <IconButton iconProps={{ iconName: 'Settings' }} title="Widget settings" ariaLabel="Widget settings" onClick={() => onOpenSettings(instance.id)} styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 13 } }} />}
              <IconButton iconProps={{ iconName: 'Delete' }} title="Remove widget" ariaLabel="Remove widget" onClick={() => removeWidget(instance.id)} styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 13 } }} />
            </>
          )}
          <IconButton
            iconProps={{ iconName: instance.collapsed ? 'ChevronDown' : 'ChevronUp' }}
            title={instance.collapsed ? 'Expand' : 'Collapse'}
            ariaLabel={instance.collapsed ? 'Expand widget' : 'Collapse widget'}
            aria-expanded={!instance.collapsed}
            onClick={() => updateWidget(instance.id, { collapsed: !instance.collapsed })}
            styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 11 } }}
          />
        </div>
      </header>
      <div className={styles.widgetBody}>{children}</div>
    </section>
  );
};
