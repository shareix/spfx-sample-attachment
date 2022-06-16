import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IHelloWorldProps {
  description: string;
  ListName: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  InputText: string;
}
