import * as React from 'react';
import { Icon, IconButton, Text, DefaultButton } from '@fluentui/react';
import { IFileResult, IMailResult, IEventResult, ISiteResult } from '../models';
import { IUnifiedSearchClassNames, fileColor, fileIconName } from './styles';

export function fmtDate(iso?: string): string {
  if (!iso) { return ''; }
  const date = new Date(iso);
  if (isNaN(date.getTime())) { return ''; }
  return date.toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' });
}

export function fmtTime(iso?: string): string {
  if (!iso) { return ''; }
  const date = new Date(iso);
  if (isNaN(date.getTime())) { return ''; }
  return date.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
}

export function fmtSize(bytes?: number): string {
  if (!bytes || bytes <= 0) { return ''; }
  if (bytes < 1024 * 1024) { return `${Math.max(1, Math.round(bytes / 1024))} KB`; }
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/**
 * Graph hit summaries arrive with <c0>…</c0> hit-highlight markers and <ddd/>
 * ellipses. Render them as React nodes — never raw HTML — so nothing from the
 * index can inject markup.
 */
export function renderSummary(summary?: string): React.ReactNode {
  if (!summary) { return null; }
  const withEllipses = summary.split(/<ddd\/>/g).join('…');
  const parts = withEllipses.split(/<\/?c0>/g);
  return parts.map((part, index) =>
    index % 2 === 1
      ? <strong key={index}>{part}</strong>
      : <React.Fragment key={index}>{part}</React.Fragment>
  );
}

interface IRowCommon {
  classNames: IUnifiedSearchClassNames;
  openInNewTab: boolean;
}

export interface IFileRowProps extends IRowCommon {
  file: IFileResult;
  enablePreview: boolean;
  onPreview: (file: IFileResult) => void;
  onShare: (file: IFileResult) => void;
  onCopyLink: (file: IFileResult) => void;
}

export const FileRow: React.FunctionComponent<IFileRowProps> = props => {
  const { file, classNames } = props;
  const open = (): void => {
    if (props.enablePreview && file.driveId) {
      props.onPreview(file);
    } else {
      window.open(file.webUrl, props.openInNewTab ? '_blank' : '_self', 'noopener');
    }
  };
  return (
    <div
      className={classNames.rowHover}
      role="link"
      tabIndex={0}
      onClick={open}
      onKeyDown={e => { if (e.key === 'Enter') { open(); } }}
    >
      <div className={classNames.fileIcon} style={{ background: fileColor(file.fileType) }}>
        <Icon iconName={fileIconName(file.fileType)} />
      </div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <Text variant="mediumPlus" block styles={{ root: { fontWeight: 600 } }}>{file.name}</Text>
        <div className={classNames.meta}>
          <span>{file.siteName}</span>
          {file.modified && <span>· Modified {fmtDate(file.modified)}{file.modifiedBy ? ` by ${file.modifiedBy}` : ''}</span>}
          {!!file.size && <span>· {fmtSize(file.size)}</span>}
          {file.isOneDrive && <span>· OneDrive</span>}
        </div>
        <p className={classNames.summary}>{renderSummary(file.summary)}</p>
      </div>
      <div style={{ display: 'flex', alignSelf: 'center' }} onClick={e => e.stopPropagation()}>
        {props.enablePreview && !!file.driveId &&
          <IconButton iconProps={{ iconName: 'View' }} title="Preview" ariaLabel="Preview"
            onClick={() => props.onPreview(file)} />}
        {!!file.driveId &&
          <IconButton iconProps={{ iconName: 'Share' }} title="Share" ariaLabel="Share"
            onClick={() => props.onShare(file)} />}
        <IconButton iconProps={{ iconName: 'Link' }} title="Copy link" ariaLabel="Copy link"
          onClick={() => props.onCopyLink(file)} />
      </div>
    </div>
  );
};

export interface IMailRowProps extends IRowCommon {
  mail: IMailResult;
}

export const MailRow: React.FunctionComponent<IMailRowProps> = props => {
  const { mail, classNames } = props;
  const open = (): void => {
    if (mail.webLink) {
      window.open(mail.webLink, props.openInNewTab ? '_blank' : '_self', 'noopener');
    }
  };
  return (
    <div
      className={classNames.rowHover}
      role="link"
      tabIndex={0}
      onClick={open}
      onKeyDown={e => { if (e.key === 'Enter') { open(); } }}
    >
      <Icon iconName="Mail" style={{ fontSize: 18, alignSelf: 'flex-start', marginTop: 3 }} />
      <div style={{ flex: 1, minWidth: 0 }}>
        <Text variant="mediumPlus" block styles={{ root: { fontWeight: 600 } }}>{mail.subject}</Text>
        <div className={classNames.meta}>
          <span>{mail.from}</span>
          {mail.hasAttachments && <Icon iconName="Attach" style={{ fontSize: 12 }} />}
        </div>
        <p className={classNames.summary}>{renderSummary(mail.preview)}</p>
      </div>
      <div className={classNames.meta} style={{ whiteSpace: 'nowrap', alignSelf: 'flex-start' }}>
        {fmtDate(mail.received)}
      </div>
    </div>
  );
};

export interface IEventRowProps extends IRowCommon {
  event: IEventResult;
}

export const EventRow: React.FunctionComponent<IEventRowProps> = props => {
  const { event, classNames } = props;
  const start = event.start ? new Date(event.start) : undefined;
  const open = (): void => {
    if (event.webLink) {
      window.open(event.webLink, props.openInNewTab ? '_blank' : '_self', 'noopener');
    }
  };
  return (
    <div
      className={classNames.rowHover}
      role="link"
      tabIndex={0}
      onClick={open}
      onKeyDown={e => { if (e.key === 'Enter') { open(); } }}
    >
      <div style={{
        width: 48, flexShrink: 0, textAlign: 'center', border: '1px solid rgba(0,0,0,.1)',
        borderRadius: 6, padding: '4px 0 6px'
      }}>
        <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase', color: '#c0392b' }}>
          {start ? start.toLocaleDateString(undefined, { month: 'short' }) : '—'}
        </div>
        <div style={{ fontSize: 19, fontWeight: 700, lineHeight: 1.1 }}>
          {start ? start.getDate() : ''}
        </div>
      </div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <Text variant="mediumPlus" block styles={{ root: { fontWeight: 600 } }}>{event.subject}</Text>
        <div className={classNames.meta}>
          {event.start && <span><Icon iconName="Clock" style={{ fontSize: 12 }} /> {fmtTime(event.start)}–{fmtTime(event.end)}</span>}
          {event.location && <span><Icon iconName={event.isOnline ? 'Video' : 'MapPin'} style={{ fontSize: 12 }} /> {event.location}</span>}
          {event.organizer && <span><Icon iconName="Contact" style={{ fontSize: 12 }} /> {event.organizer}</span>}
        </div>
      </div>
      {event.isOnline && event.webLink &&
        <DefaultButton
          text="Open"
          iconProps={{ iconName: 'TeamsLogo' }}
          styles={{ root: { alignSelf: 'center' } }}
          onClick={e => { e.stopPropagation(); open(); }}
        />}
    </div>
  );
};

export interface ISiteCardProps extends IRowCommon {
  site: ISiteResult;
}

export const SiteCard: React.FunctionComponent<ISiteCardProps> = props => {
  const { site, classNames } = props;
  const initials = site.name.split(' ').map(word => word.charAt(0)).slice(0, 2).join('').toUpperCase();
  const open = (): void => {
    window.open(site.webUrl, props.openInNewTab ? '_blank' : '_self', 'noopener');
  };
  return (
    <div
      className={classNames.card + ' ' + classNames.rowHover}
      role="link"
      tabIndex={0}
      onClick={open}
      onKeyDown={e => { if (e.key === 'Enter') { open(); } }}
      style={{ padding: 16 }}
    >
      <div className={classNames.fileIcon} style={{ background: '#038387', width: 44, height: 44, fontSize: 15, fontWeight: 700 }}>
        {initials}
      </div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <Text variant="mediumPlus" block styles={{ root: { fontWeight: 600 } }}>{site.name}</Text>
        <p className={classNames.summary}>{renderSummary(site.description)}</p>
      </div>
    </div>
  );
};
