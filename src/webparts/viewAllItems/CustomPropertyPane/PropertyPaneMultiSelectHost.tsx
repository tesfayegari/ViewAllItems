import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IItemProp, IMultiSelectProp, IMultiSelectPropInternal } from './PropertyPaneMultiSelect';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IMultiSelectHostProp extends IMultiSelectPropInternal {
    stateKey: string;
}
export interface IMultiSelectHostState {
    items?: IItemProp[];
    selectedItems: string[];
}
 
export class MultiSelectHost extends React.Component<IMultiSelectHostProp, IMultiSelectHostState>{
    constructor(props: IMultiSelectHostProp, state: IMultiSelectHostState) {
        super(props);
        SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/multiple-select/1.2.0/multiple-select.min.css");
        this.state = ({
            items: [],
            selectedItems: this.props.selectedItemIds
        });
        this.onClick = this.onClick.bind(this);        
    }

    public componentDidMount(): void {
        this.loadOptions();
      }

      public componentDidUpdate(prevProps: IMultiSelectHostProp, prevState: IMultiSelectHostState): void {
        if (this.props.disabled !== prevProps.disabled ||
            this.props.stateKey !== prevProps.stateKey) {                
                this.loadOptions();
            }        
      }
      private loadOptions(): void {
        this.setState({
            items: [],
            selectedItems: this.props.selectedItemIds
        });
    
        this.props.onload()
          .then((items) => {
            this.setState({
              items: items,
              selectedItems: this.props.selectedItemIds
            });
            this._applyMultiSelect(this.props.selectedKey, this.props.selectedItemIds);
          }, (error: any): void => {
            this.setState((prevState: IMultiSelectHostState, props: IMultiSelectProp): IMultiSelectHostState => {
              prevState.items = [];
              prevState.selectedItems = this.props.selectedItemIds;
              return prevState;
            });
          });
      }
    private _applyMultiSelect(selectControlId: string, selectedIds: string[]) {
        SPComponentLoader.loadScript("https://code.jquery.com/jquery-3.2.1.min.js", { globalExportsName: 'jQuery' })
            .then((jQuery: any): void => {
                SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/multiple-select/1.2.0/multiple-select.min.js", { globalExportsName: 'jQuery' })
                    .then(() => {
                        jQuery("#" + selectControlId + "").multipleSelect({
                            width: "100%",
                            selectAll: false,
                            onClick: (item) => { this.onClick(item); }
                        });
                        jQuery("#" + selectControlId + "").multipleSelect('setSelects', selectedIds);
                    });
            });
    }
 
    private getAllItems(): IItemProp[] {
        let resours: IItemProp[] = [];
        this.props.onload().then((items) => {
            resours = items;
        });
        //resours = this.props.options;
        return resours;
    }
    private onClick(item: any) {
        let oldValues = this.props.properties[this.props.targetProperty];
        if (item.checked) {
            this.state.selectedItems.push(item.value);
        }
        else {
            var index = this.getSelectedNodePosition(item);
            if (index != -1)
                this.state.selectedItems.splice(index, 1);
        }
        this.setState(this.state);
 
        this.props.properties[this.props.targetProperty] = this.state.selectedItems;
        this.props.onPropChange(this.props.targetProperty, oldValues, this.state.selectedItems);
    }
    private getSelectedNodePosition(node): number {
        for (var i = 0; i < this.state.selectedItems.length; i++) {
            if (node.value === this.state.selectedItems[i])
                return i;
        }
        return -1;
    }
 
    public render(): JSX.Element {
        const allItems = this.state.items !== undefined? this.state.items.map((item: IItemProp) => {
            return <option id={item.key} value={item.key}>{item.text}</option>;
          }) : <option>Select List</option>;
        return (
            <div>
                <label>{this.props.label}</label>
                <br/>
                <select id={this.props.selectedKey}>
                    {allItems}
                </select>
            </div>
        );
    }
}