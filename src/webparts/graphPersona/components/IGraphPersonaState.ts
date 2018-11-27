import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IGraphPersonaState {
	people: MicrosoftGraph.Person[];
	pictures: MicrosoftGraph.Photo[];
}
