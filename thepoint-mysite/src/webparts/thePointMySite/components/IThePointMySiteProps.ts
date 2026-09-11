import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { IThePointMySiteWebPartProps } from '../models';
import type { GraphService } from '../services/GraphService';
import type { SharePointService } from '../services/SharePointService';

export interface IThePointMySiteProps {
  context: WebPartContext;
  graph: GraphService;
  sp: SharePointService;
  theme: IReadonlyTheme | undefined;
  isDarkTheme: boolean;
  properties: IThePointMySiteWebPartProps;
  isPageEditMode: boolean;
  instanceId: string;
  onSaveAsDefault: (json: string, audience: string) => void;
}
