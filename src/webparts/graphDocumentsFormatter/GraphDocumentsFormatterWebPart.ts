import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
	BaseClientSideWebPart,
	IPropertyPaneConfiguration,
	PropertyPaneTextField,
	PropertyPaneSlider
} from '@microsoft/sp-webpart-base';

import * as strings from 'GraphDocumentsFormatterWebPartStrings';
import GraphDocumentsFormatter from './components/GraphDocumentsFormatter';
import { IGraphDocumentsFormatterProps } from './components/IGraphDocumentsFormatterProps';
import { MSGraphClient } from '@microsoft/sp-http';
export interface IGraphDocumentsFormatterWebPartProps {
	api: string;
	numberOfDocuments: number;
}
export default class GraphDocumentsFormatterWebPart extends BaseClientSideWebPart<
	IGraphDocumentsFormatterWebPartProps
> {
	public render(): void {
		this.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
			const element: React.ReactElement<
				IGraphDocumentsFormatterProps
			> = React.createElement(GraphDocumentsFormatter, {
				graphClient: client,
				numberOfDocuments: this.properties.numberOfDocuments,
				api: this.properties.api
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
					header: {
						description: strings.PropertyPaneDescription
					},
					groups: [
						{
							groupName: strings.BasicGroupName,
							groupFields: [
								PropertyPaneSlider('numberOfDocuments', {
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
