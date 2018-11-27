import * as React from 'react';
import { IGraphPersonaProps } from './IGraphPersonaProps';
import 'office-ui-fabric-core/dist/css/fabric.min.css';
import { IGraphPersonaState } from './IGraphPersonaState';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Persona, PersonaSize } from 'office-ui-fabric-react/lib/components/Persona';
import { List } from 'office-ui-fabric-react/lib/List';
import { FocusZone } from 'office-ui-fabric-react/lib/FocusZone';
import { Log } from '@microsoft/sp-core-library';
const LOG_SOURCE: string = 'Frequent Contacts Webpart';
export default class GraphPersona extends React.Component<IGraphPersonaProps, IGraphPersonaState> {
	constructor(props: IGraphPersonaProps) {
		super(props);
		this.state = {
			people: [],
			pictures: []
		};
	}

	public componentDidMount(): void {
		try {
			//API will be made configurable via Webpart property
			this.props.graphClient
				.api(
					"me/people?$filter=personType/subclass eq 'OrganizationUser'" +
						'&$top=' +
						this.props.numberOfContacts
				)
				//	.api(this.props.graphAPI + '&?$top=' + this.props.numberOfContacts)
				.get((error: any, userResponse: any, rawResponse?: any) => {
					const peopleResults: MicrosoftGraph.Person[] = userResponse.value;
					console.log('peopleResults', peopleResults);
					console.log(this.props.numberOfContacts);
					this.setState({
						people: peopleResults
					});
					//Code to put images for contacts, depends on requirement as this increases the number of API Calls
					/* peopleResults.forEach((person) => {
					this.props.graphClient
						.api('/users/' + person.userPrincipalName + '/photo/$Value')
						.get()
						.then((picture: any, pictureResponse?: any) => {
							var pictureResults: MicrosoftGraph.Photo[] = pictureResponse.value;
              console.log('picture', pictureResults);
              this.state.pictures.push(pictureResults[0]);
						});
				}); */
				});
		} catch (error) {
			Log.error(LOG_SOURCE, error);
		}
	}
	public componentDidUpdate(prevProps: IGraphPersonaProps, prevState: IGraphPersonaState, prevContext: any): void {
		try {
			//Update wepart results if its property changes
			if (
				this.props.numberOfContacts !== prevProps.numberOfContacts ||
				this.props.graphAPI !== prevProps.graphAPI
			) {
				this.props.graphClient
					.api(
						"me/people?$filter=personType/subclass eq 'OrganizationUser'&$top=" +
							this.props.numberOfContacts
					)
					//.api(this.props.graphAPI + '&?$top=' + this.props.numberOfContacts)
					.get((error: any, userResponse: any, rawResponse?: any) => {
						console.log(userResponse.value);
						const peopleResults: MicrosoftGraph.Person[] = userResponse.value;
						console.log('Property Updated');
						console.log('peopleResults', peopleResults);
						console.log(this.props.numberOfContacts);
						this.setState({
							people: peopleResults
						});
					});
			}
		} catch (error) {
			Log.error(LOG_SOURCE, error);
		}
	}
	private _onRenderEventCell(item: MicrosoftGraph.Person, index: number | undefined): JSX.Element {
		return (
			<div className="ms-Grid-col ms-sm4 ms-md4 ms-lg4">
				<a href={item.imAddress} style={{ 'text-decoration': 'none' }}>
					<Persona text={item.displayName} size={PersonaSize.large} />
				</a>
			</div>
		);
	}
	public render(): React.ReactElement<IGraphPersonaProps> {
		return (
			<div className="ms-Grid" dir="ltr">
				<div className="ms-Grid-row">
					<FocusZone>
						<List items={this.state.people} onRenderCell={this._onRenderEventCell} />
					</FocusZone>
				</div>
			</div>
		);
	}
}
