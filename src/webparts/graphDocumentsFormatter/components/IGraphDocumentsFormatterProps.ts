import { MSGraphClient } from '@microsoft/sp-http';

export interface IGraphDocumentsFormatterProps {
  graphClient: MSGraphClient;
  api: string;
  numberOfDocuments: number;
}
