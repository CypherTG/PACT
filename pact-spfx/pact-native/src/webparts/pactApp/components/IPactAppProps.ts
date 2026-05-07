import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPactAppProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  viewMode: 'Admin' | 'PublicReport';
}
