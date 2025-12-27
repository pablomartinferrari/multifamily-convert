import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IMultifamilyProcessorProps {
  description: string;
  context: WebPartContext;
  siteUrl: string;
  listId: string;
  azureFunctionUrl: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
