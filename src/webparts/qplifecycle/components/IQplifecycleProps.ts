import { SPHttpClient } from '@microsoft/sp-http';

export interface IQpcycleProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  listName: string;
}
