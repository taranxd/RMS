import * as React from 'react';
import { IGraphEventsListProps } from './IGraphEventsListProps';
import { IGraphEventsListState } from './IGraphEventsListState';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { List } from 'office-ui-fabric-react/lib/List';
import { format } from 'date-fns';
export default class GraphEventsList extends React.Component<IGraphEventsListProps, IGraphEventsListState> {
	constructor(props: IGraphEventsListProps) {
		super(props);

		this.state = {
			events: []
		};
	}

	public componentDidMount(): void {
		let futureDate = new Date();
		futureDate.setDate(futureDate.getDate() + this.props.eventDays);
		let api =
			'/me/calendar/calendarView?startDateTime=' +
			new Date().toISOString() +
			'&endDateTime=' +
			futureDate.toISOString();
		this.props.graphClient
			.api(api)
			// .api('/me/events')
			.get((error: any, eventsResponse: any, rawResponse?: any) => {
				const calendarEvents: MicrosoftGraph.Event[] = eventsResponse.value;
				console.log('calendarEvents', calendarEvents);
				this.setState({ events: calendarEvents });
			});
	}
	public componentDidUpdate(
		prevProps: IGraphEventsListProps,
		prevState: IGraphEventsListState,
		prevContext: any
	): void {
		if (this.props.eventDays !== prevProps.eventDays) {
			let futureDate = new Date();
			futureDate.setDate(futureDate.getDate() + this.props.eventDays);
			let api =
				'/me/calendar/calendarView?startDateTime=' +
				new Date().toISOString() +
				'&endDateTime=' +
				futureDate.toISOString();
			this.props.graphClient.api(api).get((error: any, eventsResponse: any, rawResponse?: any) => {
				const calendarEvents: MicrosoftGraph.Event[] = eventsResponse.value;
				console.log('calendarEvents', calendarEvents);
				this.setState({ events: calendarEvents });
			});
		}
	}
	private _onRenderEventCell(item: MicrosoftGraph.Event, index: number | undefined): JSX.Element {
		return (
			<div>
				<h3>{item.subject}</h3>
				{format(new Date(item.start.dateTime), 'MMMM Mo, YYYY h:mm A')} -{' '}
				{format(new Date(item.end.dateTime), 'h:mm A')}
			</div>
		);
	}

	public render(): React.ReactElement<IGraphEventsListProps> {
		return <List items={this.state.events} onRenderCell={this._onRenderEventCell} />;
	}
}
