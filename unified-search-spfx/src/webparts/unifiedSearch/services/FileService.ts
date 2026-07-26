import { MSGraphClientV3, MSGraphClientFactory } from '@microsoft/sp-http';
import { IFileResult, IEventResult } from '../models';

export class FileService {

  private _clientPromise: Promise<MSGraphClientV3>;

  constructor(factory: MSGraphClientFactory) {
    this._clientPromise = factory.getClient('3');
  }

  /**
   * Embeddable, short-lived preview URL from the Graph preview endpoint.
   * Works for Office documents, PDF and 300+ other formats.
   */
  public async getPreviewUrl(file: IFileResult): Promise<string | undefined> {
    if (!file.driveId || !file.itemId) {
      return undefined;
    }
    const client = await this._clientPromise;
    const response: { getUrl?: string } = await client
      .api(`/drives/${file.driveId}/items/${file.itemId}/preview`)
      .post({});
    return response.getUrl;
  }

  /** Organisation-scoped view link ("people in your org with the link can view"). */
  public async createViewLink(file: IFileResult): Promise<string> {
    if (!file.driveId || !file.itemId) {
      return file.webUrl;
    }
    const client = await this._clientPromise;
    const response: { link?: { webUrl?: string } } = await client
      .api(`/drives/${file.driveId}/items/${file.itemId}/createLink`)
      .post({ type: 'view', scope: 'organization' });
    return (response.link && response.link.webUrl) || file.webUrl;
  }

  /** Direct sharing invitation, mirroring the OneDrive/SharePoint share dialog. */
  public async share(
    file: IFileResult, recipients: string[], canEdit: boolean, message: string
  ): Promise<void> {
    if (!file.driveId || !file.itemId) {
      throw new Error('This item cannot be shared from here.');
    }
    const client = await this._clientPromise;
    await client
      .api(`/drives/${file.driveId}/items/${file.itemId}/invite`)
      .post({
        recipients: recipients.map(email => ({ email })),
        requireSignIn: true,
        sendInvitation: true,
        roles: [canEdit ? 'write' : 'read'],
        message: message || ''
      });
  }

  /** Recently used files for the zero-query state. */
  public async getRecentFiles(top: number): Promise<IFileResult[]> {
    const client = await this._clientPromise;
    const response: {
      value?: {
        id?: string;
        resourceVisualization?: { title?: string; type?: string; containerDisplayName?: string };
        resourceReference?: { webUrl?: string };
        lastUsed?: { lastAccessedDateTime?: string };
      }[];
    } = await client.api('/me/insights/used').top(top).get();
    return (response.value || []).map(insight => {
      const visual = insight.resourceVisualization || {};
      const reference = insight.resourceReference || {};
      return {
        key: insight.id || reference.webUrl || Math.random().toString(36),
        name: visual.title || '',
        fileType: (visual.type || '').toLowerCase(),
        webUrl: reference.webUrl || '',
        siteName: visual.containerDisplayName || '',
        modified: insight.lastUsed ? insight.lastUsed.lastAccessedDateTime : undefined,
        isOneDrive: (reference.webUrl || '').indexOf('-my.sharepoint.com') !== -1
      } as IFileResult;
    }).filter(file => !!file.name && !!file.webUrl);
  }

  /** Next few calendar entries for the zero-query state. */
  public async getUpcomingEvents(top: number): Promise<IEventResult[]> {
    const client = await this._clientPromise;
    const now = new Date();
    const inTwoWeeks = new Date(now.getTime() + 14 * 86400000);
    const response: {
      value?: {
        id?: string; subject?: string; webLink?: string; isOnlineMeeting?: boolean;
        start?: { dateTime?: string }; end?: { dateTime?: string };
        location?: { displayName?: string };
        organizer?: { emailAddress?: { name?: string } };
      }[];
    } = await client
      .api('/me/calendarview')
      .query({ startDateTime: now.toISOString(), endDateTime: inTwoWeeks.toISOString() })
      .select('id,subject,start,end,location,organizer,isOnlineMeeting,webLink')
      .orderby('start/dateTime')
      .top(top)
      .get();
    return (response.value || []).map(event => ({
      key: event.id || Math.random().toString(36),
      subject: event.subject || '(no subject)',
      start: event.start ? event.start.dateTime : undefined,
      end: event.end ? event.end.dateTime : undefined,
      location: event.location ? event.location.displayName : undefined,
      organizer: event.organizer && event.organizer.emailAddress ? event.organizer.emailAddress.name : undefined,
      isOnline: !!event.isOnlineMeeting,
      webLink: event.webLink
    }));
  }
}
