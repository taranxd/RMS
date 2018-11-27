import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneSlider } from '@microsoft/sp-webpart-base';

import * as strings from 'GraphEventsListWebPartStrings';
import GraphEventsList from './components/graphEventsList';
import { IGraphEventsListProps } from './components/IGraphEventsListProps';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IGraphEventsListWebPartProps {
	eventDays: number;
}

export default class GraphEventsListWebPart extends BaseClientSideWebPart<IGraphEventsListWebPartProps> {
	public render(): void {
		this.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
			const element: React.ReactElement<IGraphEventsListProps> = React.createElement(GraphEventsList, {
				graphClient: client,
				eventDays: this.properties.eventDays
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
								PropertyPaneSlider('eventDays', {
									label: 'Duration in Days',
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
