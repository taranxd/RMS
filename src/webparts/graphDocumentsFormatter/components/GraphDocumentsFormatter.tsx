import * as React from 'react';
import { IGraphDocumentsFormatterProps } from './IGraphDocumentsFormatterProps';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import {
	DocumentCard,
	DocumentCardActivity,
	DocumentCardTitle,
	DocumentCardType
} from 'office-ui-fabric-react/lib/DocumentCard';
import { List } from 'office-ui-fabric-react/lib/List';
import { FocusZone } from 'office-ui-fabric-react/lib/FocusZone';
import { IGraphDocumentsState } from './IGraphDocumentsState';
import { Log } from '@microsoft/sp-core-library';
const LOG_SOURCE: string = 'Important Documents Webpart';
export default class GraphPersona extends React.Component<IGraphDocumentsFormatterProps, IGraphDocumentsState> {
	constructor(props: IGraphDocumentsFormatterProps) {
		super(props);
		this.state = {
			Documents: []
		};
	}
	public componentDidMount(): void {
		try {
			var tempAttachments = [];
			this.props.graphClient
				.api(
					'me/mailFolders/inbox/messages?$filter=hasAttachments eq true&$top=' + this.props.numberOfDocuments
				)
				.get((error: any, apiResponse: any, rawResponse?: any) => {
					const apiResults: MicrosoftGraph.Message[] = apiResponse.value;
					var count = 0;
					apiResults.forEach((message) => {
						this.props.graphClient
							.api('me/messages/' + message.id + '/attachments')
							.get()
							.then((attachmentAPIResponse: any, attachmentResponse?: any) => {
								var attachmentResults: MicrosoftGraph.FileAttachment[] = attachmentAPIResponse.value;
								console.log('attachments', attachmentResults);
								for (var i = 0; i < attachmentResults.length; i++) {
									if (typeof tempAttachments !== 'undefined') {
										if (tempAttachments.length < this.props.numberOfDocuments) {
											tempAttachments.push(attachmentResults[i]);
											console.log('length', tempAttachments.length);
											this.setState({
												Documents: tempAttachments
											});
										} else {
											if (tempAttachments.length == this.props.numberOfDocuments)
												this.setState({
													Documents: tempAttachments
												});
											else break;
										}
									}
								}
							});
						// attachmentResults.forEach((attachment) => {
						// 	/* if (typeof this.state.Documents !== undefined) {
						// 			if (this.state.Documents.length < this.props.numberOfDocuments) {
						// 				this.state.Documents.push(attachment);
						// 				this.setState({
						// 					Documents: this.state.Documents
						// 				});
						// 				console.log('length', this.state.Documents.length);
						// 			}
						// 		} else {
						// 			this.state.Documents.push(attachment);
						//     } */
						// 	if (typeof tempAttachments !== 'undefined') {
						// 		if (tempAttachments.length < this.props.numberOfDocuments) {
						// 			tempAttachments.push(attachment);
						// 			console.log('length', tempAttachments.length);
						// 		} else {
						// 			this.setState({
						// 				Documents: tempAttachments
						// 			});
						// 		}
						// 	}
						// });
						//	});
					});
				});
		} catch (error) {
			Log.error(LOG_SOURCE, error);
		}
	}
	public componentDidUpdate(
		prevProps: IGraphDocumentsFormatterProps,
		prevState: IGraphDocumentsState,
		prevContext: any
	): void {
		try {
			if (this.props.numberOfDocuments !== prevProps.numberOfDocuments) {
				var tempAttachments = [];
				this.props.graphClient
					.api(
						'me/mailFolders/inbox/messages?$filter=hasAttachments eq true&$top=' +
							this.props.numberOfDocuments
					)
					.get((error: any, apiResponse: any, rawResponse?: any) => {
						const apiResults: MicrosoftGraph.Message[] = apiResponse.value;
						console.log('Updated');
						console.log('Mails', apiResults);
						console.log(this.props.numberOfDocuments);
						var count = 0;
						apiResults.forEach((message) => {
							this.props.graphClient
								.api('me/messages/' + message.id + '/attachments')
								.get()
								.then((attachmentAPIResponse: any, attachmentResponse?: any) => {
									var attachmentResults: MicrosoftGraph.FileAttachment[] =
										attachmentAPIResponse.value;
									console.log('attachments', attachmentResults);
									for (var i = 0; i < attachmentResults.length; i++) {
										if (typeof tempAttachments !== 'undefined') {
											if (tempAttachments.length < this.props.numberOfDocuments) {
												tempAttachments.push(attachmentResults[i]);
												console.log('length', tempAttachments.length);
												this.setState({
													Documents: tempAttachments
												});
											} else {
												if (tempAttachments.length == this.props.numberOfDocuments)
													this.setState({
														Documents: tempAttachments
													});
												else break;
											}
										}
									}
								});
							// attachmentResults.forEach((attachment) => {
							// 	/* if (typeof this.state.Documents !== undefined) {
							// 			if (this.state.Documents.length < this.props.numberOfDocuments) {
							// 				this.state.Documents.push(attachment);
							// 				this.setState({
							// 					Documents: this.state.Documents
							// 				});
							// 				console.log('length', this.state.Documents.length);
							// 			}
							// 		} else {
							// 			this.state.Documents.push(attachment);
							//     } */
							// 	if (typeof tempAttachments !== 'undefined') {
							// 		if (tempAttachments.length < this.props.numberOfDocuments) {
							// 			tempAttachments.push(attachment);
							// 			console.log('length', tempAttachments.length);
							// 		} else {
							// 			this.setState({
							// 				Documents: tempAttachments
							// 			});
							// 		}
							// 	}
							// });
							//	});
						});
					});
			}
		} catch (error) {
			Log.error(LOG_SOURCE, error);
		}
	}
	private _onRenderEventCell(item: MicrosoftGraph.FileAttachment, index: number | undefined): JSX.Element {
		console.log(item);
		return (
			<div className="ms-Grid-col ms-sm4 ms-md4 ms-lg4">
				<a
					data-content={item.contentBytes}
					data-fileType={item.contentType}
					data-fileName={item.name}
					style={{ 'text-decoration': 'none' }}
					onClick={(e) => {
						console.log(e.currentTarget.getAttribute('data-content'));
						const url = window.URL;
						let decodedString = atob(e.currentTarget.getAttribute('data-content'));
						let bytes = new Uint8Array(decodedString.length);
						for (var i = 0; i < decodedString.length; i++) {
							bytes[i] = decodedString.charCodeAt(i);
						}
						const blob = new Blob([ bytes ], { type: e.currentTarget.getAttribute('data-fileType') });
						const blobUrl = url.createObjectURL(blob);
						let a = document.createElement('a');
						document.body.appendChild(a);
						a.href = blobUrl;
						a.download = e.currentTarget.getAttribute('data-fileName');
						a.click();
						window.URL.revokeObjectURL(blobUrl);
					}}
				>
					<DocumentCard type={DocumentCardType.normal}>
						<DocumentCardTitle title={item.name} shouldTruncate={true} />
						<DocumentCardActivity
							activity={new Date(item.lastModifiedDateTime).toDateString()}
							people={[ { name: '', profileImageSrc: '', initials: '' } ]}
						/>
					</DocumentCard>
				</a>
			</div>
		);
	}
	public render(): React.ReactElement<IGraphDocumentsFormatterProps> {
		return (
			<div className="ms-Grid" dir="ltr">
				<div className="ms-Grid-row">
					<FocusZone>
						<List items={this.state.Documents} onRenderCell={this._onRenderEventCell} />
					</FocusZone>
				</div>
			</div>
		);
	}
}
