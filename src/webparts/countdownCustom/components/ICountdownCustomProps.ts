import { SPHttpClient } from '@microsoft/sp-http';

export interface ICountdownCustomWebPartProps {
  listTitle: string;
  label: string;
  buttonText: string;
  description: string;
  expiredMessage: string;
  backgroundImageUrl: string;
  overlayOpacity: number;
  backgroundImage?: string;
}

export interface ICountdownCustomProps extends ICountdownCustomWebPartProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  isEditMode: boolean;
}
