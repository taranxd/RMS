import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
	BaseClientSideWebPart,
	IPropertyPaneConfiguration,
	PropertyPaneTextField,
	PropertyPaneSlider
} from '@microsoft/sp-webpart-base';
import * as strings from 'GraphPersonaWebPartStrings';
import GraphPersona from './components/GraphPersona';
import { IGraphPersonaProps } from './components/IGraphPersonaProps';

import { MSGraphClient } from '@microsoft/sp-http';
export interface IGraphPersonaWebPartProps {
	API: string;
	numberOfContacts: number;
}
export default class GraphPersonaWebPart extends BaseClientSideWebPart<IGraphPersonaWebPartProps> {
	public render(): void {
		this.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
			const element: React.ReactElement<IGraphPersonaProps> = React.createElement(GraphPersona, {
				graphClient: client,
				numberOfContacts: this.properties.numberOfContacts,
				graphAPI: this.properties.API
			});

			ReactDom.render(element, this.domElement);
		});
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					groups: [
						{
							groupName: strings.BasicGroupName,
							groupFields: [
								/* PropertyPaneTextField('API', {
									label: 'Graph API',
									value: "me/people?$filter=personType/subclass eq 'OrganizationUser'"
								}), */
								PropertyPaneSlider('numberOfContacts', {
									label: 'Max Items',
									min: 1,
									max: 10,
									value: 1,
									showValue: true,
									step: 1
								})
							]
						}
					]
				}
			]
		};
	}
}
