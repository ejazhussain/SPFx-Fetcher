import { SPHttpClientConfiguration } from '@microsoft/sp-http';
export interface ISPFxFetcher {
   get: (endpoint: string, config?: SPHttpClientConfiguration) => Promise<any>;
}