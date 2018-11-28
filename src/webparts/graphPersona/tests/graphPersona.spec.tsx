/// <reference types="mocha" />

import * as React from 'react';
import * as Adapter from 'enzyme-adapter-react-15';
import * as Sinon from 'sinon';
import { assert, expect } from 'chai';
import { configure, mount, ReactWrapper } from 'enzyme';
import GraphPersona from '../components/GraphPersona';
import { IGraphPersonaState } from '../components/IGraphPersonaState';
import { IGraphPersonaProps } from '../components/IGraphPersonaProps';
import { MSGraphClient } from '@microsoft/sp-http';
import 'office-ui-fabric-core/dist/css/fabric.min.css';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Persona, PersonaSize } from 'office-ui-fabric-react/lib/components/Persona';
import { List } from 'office-ui-fabric-react/lib/List';
import { FocusZone } from 'office-ui-fabric-react/lib/FocusZone';
import { Log } from '@microsoft/sp-core-library';
configure({
	adapter: new Adapter()
});

describe('<GraphPersona />', () => {
	const itemCount = 4;
	const descTxt = 'This is a Test';
	let graphClientTest = MSGraphClient;
	let componentDidMountSpy: Sinon.SinonSpy;
	let renderedElement: ReactWrapper<IGraphPersonaProps, IGraphPersonaState>;

	beforeEach(() => {
		componentDidMountSpy = Sinon.spy(GraphPersona.prototype, 'componentDidMount');
		renderedElement = mount(<GraphPersona numberOfContacts={itemCount} />);
	});

	afterEach(() => {
		renderedElement.unmount();
		componentDidMountSpy.restore();
	});

	// Test for checking if it is working
	it('Should do something', () => {
		assert.ok(true);
	});

	it('<GraphPersona /> should render something', () => {
		expect(renderedElement.find('div')).to.be.exist;
	});

	it('<GraphPersona /> should render the Focus Zone', () => {
		expect(renderedElement.find('.ms-FocusZone').text()).to.be.equals(descTxt);
	});

	it('<GraphPersona /> should render an ms-Persona', () => {
		expect(renderedElement.find('.ms-Persona')).to.be.exist;
	});

	it('<GraphPersona /> state results should not be null', () => {
		expect(renderedElement.state('people')).to.not.be.null;
	});

	it('<GraphPersona /> should call componentDidMount only once', () => {
		// Check if the componentDidMount is called once
		expect(componentDidMountSpy.calledOnce).to.equal(true);
	});

	it('<GraphPersona /> should render an ms-Persona with 3 items (using the mock data)', (done) => {
		// New instance should be created for this test due to setTimeout
		// If the global renderedElement used, the result of "ul li"" will be 10 instead of 3
		// because the state changes to 10 in the last test and
		// the last test is executed before this one bacause of setTimeout
		let renderedElement1 = mount(<GraphPersona numberOfContacts={itemCount} />);
		// Wait for 1 second to check if your mock results are retrieved
		setTimeout(() => {
			// Trigger state update
			renderedElement1.update();
			expect(renderedElement1.state('people')).to.not.be.null;
			expect(renderedElement1.find('.ms-Persona').length).to.be.equal(3);
			done(); // done is required by mocha, otherwise the test will yield SUCCESS no matter of the expect cases
		}, 1000);
	});

	it('<GraphPersona /> should render 10 list items (triggering setState from the test)', () => {
		renderedElement.setState({
			people: {
				value: [
					{ displayName: 'Mock List 1', imAddress: '1' },
					{ displayName: 'Mock List 2', imAddress: '2' },
					{ displayName: 'Mock List 3', imAddress: '3' },
					{ displayName: 'Mock List 4', imAddress: '4' },
					{ displayName: 'Mock List 5', imAddress: '5' },
					{ displayName: 'Mock List 6', imAddress: '6' },
					{ displayName: 'Mock List 7', imAddress: '7' },
					{ displayName: 'Mock List 8', imAddress: '8' },
					{ displayName: 'Mock List 9', imAddress: '9' },
					{ displayName: 'Mock List 10', imAddress: '10' }
				]
			}
		});
		expect(renderedElement.update().state('people')).to.not.be.null;
		expect(renderedElement.update().find('.ms-Persona').length).to.be.equal(10);
	});
});
