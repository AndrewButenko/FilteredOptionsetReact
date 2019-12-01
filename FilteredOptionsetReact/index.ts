import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { IDropdownOption } from "office-ui-fabric-react";
import FilteredOptionsetControl from "./FilteredOptionsetControl";
import * as React from 'react';
import * as ReactDOM from 'react-dom';

export class FilteredOptionsetReact implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private availableValues: number[];
	private container: HTMLDivElement;
	private currentValue: number | null;
	private notifyOutputChanged: () => void;

	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		let availableValuesString = context.parameters.availableValues.raw;

		if (availableValuesString == null) {
			container.innerHTML = "Property 'Available Values' is blank, configure it please for correct work";
			return;
		}

		if (availableValuesString == "val") {
			availableValuesString = "1|2";
		}

		this.availableValues = availableValuesString.split("|").map(t => parseInt(t));

		this.container = container;
		this.notifyOutputChanged = notifyOutputChanged;

		this.renderControl(context);
	}

	private renderControl(context: ComponentFramework.Context<IInputs>) {
		const currentValue = context.parameters.value.raw;

		let options = context.parameters.value.attributes!.Options.filter(option => this.availableValues.some(o => o === option.Value)).map(option =>
			({
				key: option.Value,
				text: option.Label
			}));

		let dropDownProperties = {
			options: options,
			selectedKey: currentValue,
			onSelectedChanged: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
				this.currentValue = option == null ? null : <number>option.key;
				this.notifyOutputChanged();
			}
		};

		ReactDOM.render(React.createElement(FilteredOptionsetControl, dropDownProperties), this.container);
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		this.renderControl(context);
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {
			value: this.currentValue == null ? undefined : this.currentValue
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		ReactDOM.unmountComponentAtNode(this.container);
	}
}