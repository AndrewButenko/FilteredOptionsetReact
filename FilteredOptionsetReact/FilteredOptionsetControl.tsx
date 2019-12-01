import * as React from 'react';
import {Dropdown, IDropdownOption, initializeIcons} from 'office-ui-fabric-react';

interface IFilteredOptionsetProperties {
    options: IDropdownOption[];
    selectedKey: number | null;
    onSelectedChanged: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => void;
}

initializeIcons();

export default class FilteredOptionsetControl extends React.Component<IFilteredOptionsetProperties, {}> {
    render() {
        return (
            <Dropdown
                placeHolder="--Select--"
                options={this.props.options}
                selectedKey={this.props.selectedKey}
                onChange={this.props.onSelectedChanged}
            />
        );
    }
}