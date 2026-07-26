import * as React from 'react';
import {
  Panel, PanelType, Spinner, SpinnerSize, PrimaryButton, DefaultButton,
  Dialog, DialogType, DialogFooter, TextField, Dropdown, IDropdownOption,
  MessageBar, MessageBarType, Text, IconButton
} from '@fluentui/react';
import { IFileResult } from '../models';
import { FileService } from '../services/FileService';
import { fmtDate, fmtSize } from './ResultRows';

export interface IPreviewPanelProps {
  file: IFileResult;
  files: IFileResult[];
  fileService: FileService;
  openInNewTab: boolean;
  onDismiss: () => void;
  onNavigate: (file: IFileResult) => void;
  onCopyLink: (file: IFileResult) => void;
}

export const PreviewPanel: React.FunctionComponent<IPreviewPanelProps> = props => {
  const { file } = props;
  const [previewUrl, setPreviewUrl] = React.useState<string | undefined>(undefined);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [shareOpen, setShareOpen] = React.useState<boolean>(false);

  React.useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(undefined);
    setPreviewUrl(undefined);
    props.fileService.getPreviewUrl(file)
      .then(url => {
        if (!cancelled) {
          setPreviewUrl(url);
          setLoading(false);
          if (!url) { setError('No preview is available for this file.'); }
        }
      })
      .catch(() => {
        if (!cancelled) {
          setLoading(false);
          setError('The preview could not be loaded. You can still open the file directly.');
        }
      });
    return () => { cancelled = true; };
  }, [file.key]);

  const index = props.files.map(f => f.key).indexOf(file.key);
  const previous = index > 0 ? props.files[index - 1] : undefined;
  const next = index >= 0 && index < props.files.length - 1 ? props.files[index + 1] : undefined;

  return (
    <Panel
      isOpen
      onDismiss={props.onDismiss}
      type={PanelType.large}
      headerText={`${file.name}${file.fileType ? '.' + file.fileType : ''}`}
      closeButtonAriaLabel="Close preview"
      isLightDismiss
    >
      <div style={{ display: 'flex', gap: 8, margin: '6px 0 12px', flexWrap: 'wrap' }}>
        <PrimaryButton
          iconProps={{ iconName: 'OpenInNewWindow' }}
          text="Open"
          onClick={() => window.open(file.webUrl, props.openInNewTab ? '_blank' : '_self', 'noopener')}
        />
        {!!file.driveId &&
          <DefaultButton iconProps={{ iconName: 'Share' }} text="Share" onClick={() => setShareOpen(true)} />}
        <DefaultButton iconProps={{ iconName: 'Link' }} text="Copy link" onClick={() => props.onCopyLink(file)} />
        <div style={{ flex: 1 }} />
        <IconButton iconProps={{ iconName: 'ChevronLeft' }} title="Previous result"
          disabled={!previous} onClick={() => previous && props.onNavigate(previous)} />
        <IconButton iconProps={{ iconName: 'ChevronRight' }} title="Next result"
          disabled={!next} onClick={() => next && props.onNavigate(next)} />
      </div>

      <div style={{ fontSize: 12.5, marginBottom: 10, opacity: .8 }}>
        <Text variant="small">
          {file.siteName}
          {file.modified ? ` · Modified ${fmtDate(file.modified)}` : ''}
          {file.modifiedBy ? ` by ${file.modifiedBy}` : ''}
          {file.size ? ` · ${fmtSize(file.size)}` : ''}
        </Text>
      </div>

      {loading &&
        <div style={{ padding: 60, textAlign: 'center' }}>
          <Spinner size={SpinnerSize.large} label="Loading preview…" />
        </div>}
      {!loading && error &&
        <MessageBar messageBarType={MessageBarType.warning}>{error}</MessageBar>}
      {!loading && previewUrl &&
        <iframe
          src={previewUrl}
          title={`Preview of ${file.name}`}
          style={{ width: '100%', height: 'calc(100vh - 260px)', minHeight: 420, border: '1px solid rgba(0,0,0,.1)', borderRadius: 6 }}
        />}

      {shareOpen &&
        <ShareDialog
          file={file}
          fileService={props.fileService}
          onDismiss={() => setShareOpen(false)}
        />}
    </Panel>
  );
};

export interface IShareDialogProps {
  file: IFileResult;
  fileService: FileService;
  onDismiss: () => void;
}

export const ShareDialog: React.FunctionComponent<IShareDialogProps> = props => {
  const [recipients, setRecipients] = React.useState<string>('');
  const [role, setRole] = React.useState<string>('read');
  const [message, setMessage] = React.useState<string>('');
  const [busy, setBusy] = React.useState<boolean>(false);
  const [result, setResult] = React.useState<{ ok: boolean; text: string } | undefined>(undefined);

  const roleOptions: IDropdownOption[] = [
    { key: 'read', text: 'Can view' },
    { key: 'write', text: 'Can edit' }
  ];

  const send = (): void => {
    const emails = recipients.split(/[,;\s]+/).map(email => email.trim()).filter(email => email.indexOf('@') > 0);
    if (emails.length === 0) {
      setResult({ ok: false, text: 'Enter at least one email address.' });
      return;
    }
    setBusy(true);
    setResult(undefined);
    props.fileService.share(props.file, emails, role === 'write', message)
      .then(() => {
        setBusy(false);
        setResult({ ok: true, text: `Shared with ${emails.join(', ')}.` });
        setRecipients('');
      })
      .catch(() => {
        setBusy(false);
        setResult({
          ok: false,
          text: 'Sharing failed. You may not have permission to share this item, or sharing may be restricted by policy.'
        });
      });
  };

  return (
    <Dialog
      hidden={false}
      onDismiss={props.onDismiss}
      dialogContentProps={{
        type: DialogType.normal,
        title: `Share "${props.file.name}"`
      }}
      modalProps={{ isBlocking: false }}
      minWidth={420}
    >
      <TextField
        label="Send to"
        placeholder="name@company.com, separated by commas"
        value={recipients}
        onChange={(_, value) => setRecipients(value || '')}
      />
      <Dropdown
        label="They can"
        options={roleOptions}
        selectedKey={role}
        onChange={(_, option) => option && setRole(String(option.key))}
      />
      <TextField
        label="Message (optional)"
        multiline
        rows={2}
        value={message}
        onChange={(_, value) => setMessage(value || '')}
      />
      {result &&
        <MessageBar
          messageBarType={result.ok ? MessageBarType.success : MessageBarType.error}
          styles={{ root: { marginTop: 10 } }}
        >
          {result.text}
        </MessageBar>}
      <DialogFooter>
        <DefaultButton text="Close" onClick={props.onDismiss} disabled={busy} />
        <PrimaryButton text="Send" iconProps={{ iconName: 'Send' }} onClick={send} disabled={busy} />
      </DialogFooter>
    </Dialog>
  );
};
