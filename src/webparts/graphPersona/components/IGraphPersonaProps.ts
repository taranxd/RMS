import { MSGraphClient } from '@microsoft/sp-http';

export interface IGraphPersonaProps {
	graphClient: MSGraphClient;
	numberOfContacts: number;
	graphAPI: string;
}
