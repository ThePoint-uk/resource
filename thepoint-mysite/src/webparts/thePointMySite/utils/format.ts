/** Formatting helpers shared by widgets. */

export function formatRelative(iso: string | undefined): string {
  if (!iso) {
    return '';
  }
  const date: Date = new Date(iso);
  const now: Date = new Date();
  const diffMs: number = now.getTime() - date.getTime();
  const minutes: number = Math.round(diffMs / 60000);
  if (minutes < 1) {
    return 'just now';
  }
  if (minutes < 60) {
    return `${minutes} min ago`;
  }
  const hours: number = Math.round(minutes / 60);
  if (hours < 24) {
    return `${hours} hr${hours === 1 ? '' : 's'} ago`;
  }
  const days: number = Math.round(hours / 24);
  if (days < 7) {
    return days === 1 ? 'yesterday' : `${days} days ago`;
  }
  if (days < 30) {
    const weeks: number = Math.round(days / 7);
    return `${weeks} wk${weeks === 1 ? '' : 's'} ago`;
  }
  return date.toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: date.getFullYear() === now.getFullYear() ? undefined : 'numeric' });
}

export function formatDate(iso: string | Date | undefined, options?: Intl.DateTimeFormatOptions): string {
  if (!iso) {
    return '';
  }
  const date: Date = typeof iso === 'string' ? new Date(iso) : iso;
  return date.toLocaleDateString(undefined, options || { weekday: 'short', day: 'numeric', month: 'short' });
}

export function formatTime(iso: string | Date | undefined): string {
  if (!iso) {
    return '';
  }
  const date: Date = typeof iso === 'string' ? new Date(iso) : iso;
  return date.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
}

export function formatBytes(bytes: number | undefined): string {
  if (bytes === undefined || bytes === null || isNaN(bytes)) {
    return '';
  }
  const units: string[] = ['B', 'KB', 'MB', 'GB', 'TB'];
  let value: number = bytes;
  let unit: number = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit++;
  }
  return `${value < 10 && unit > 0 ? value.toFixed(1) : Math.round(value)} ${units[unit]}`;
}

/** Fluent UI icon name for a file by extension or Graph resource type. */
export function fileIcon(nameOrType: string | undefined): string {
  const value: string = (nameOrType || '').toLowerCase();
  const ext: string = value.indexOf('.') >= 0 ? value.split('.').pop() || '' : value;
  switch (ext) {
    case 'docx': case 'doc': case 'word': return 'WordDocument';
    case 'xlsx': case 'xls': case 'csv': case 'excel': return 'ExcelDocument';
    case 'pptx': case 'ppt': case 'powerpoint': return 'PowerPointDocument';
    case 'pdf': return 'PDF';
    case 'one': case 'onenote': case 'onetoc2': return 'OneNoteLogo';
    case 'txt': case 'md': case 'text': return 'TextDocument';
    case 'png': case 'jpg': case 'jpeg': case 'gif': case 'svg': case 'bmp': case 'image': return 'Photo2';
    case 'mp4': case 'mov': case 'avi': case 'video': return 'Video';
    case 'mp3': case 'wav': case 'audio': return 'MusicNote';
    case 'zip': case 'rar': case '7z': return 'ZipFolder';
    case 'vsdx': case 'visio': return 'VisioDocument';
    case 'msg': case 'eml': return 'Mail';
    case 'aspx': case 'html': case 'htm': case 'web': case 'spsite': return 'Globe';
    case 'loop': case 'fluid': return 'Sync';
    case 'folder': return 'FabricFolder';
    default: return 'Page';
  }
}

export function fileIconColor(iconName: string): string | undefined {
  switch (iconName) {
    case 'WordDocument': return '#185ABD';
    case 'ExcelDocument': return '#107C41';
    case 'PowerPointDocument': return '#C43E1C';
    case 'PDF': return '#D13438';
    case 'OneNoteLogo': return '#7719AA';
    case 'VisioDocument': return '#3955A3';
    case 'FabricFolder': return '#DCB67A';
    default: return undefined;
  }
}

export function initials(name: string | undefined): string {
  if (!name) {
    return '?';
  }
  const parts: string[] = name.trim().split(/\s+/);
  return (parts[0]?.[0] || '') + (parts.length > 1 ? parts[parts.length - 1][0] : '');
}

export function greetingForTime(date: Date = new Date()): string {
  const hour: number = date.getHours();
  if (hour < 12) {
    return 'Good morning';
  }
  if (hour < 17) {
    return 'Good afternoon';
  }
  return 'Good evening';
}

export function uid(prefix: string = 'w'): string {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

export function truncate(text: string | undefined, max: number): string {
  if (!text) {
    return '';
  }
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

export function stripHtml(html: string | undefined): string {
  if (!html) {
    return '';
  }
  const div: HTMLDivElement = document.createElement('div');
  div.innerHTML = html;
  return div.textContent || div.innerText || '';
}
