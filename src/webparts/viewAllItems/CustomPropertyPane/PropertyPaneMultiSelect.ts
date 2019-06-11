import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IPropertyPaneCustomFieldProps, IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-webpart-base';
import { IMultiSelectHostProp, MultiSelectHost } from '../CustomPropertyPane/PropertyPaneMultiSelectHost';
 
export interface IItemProp {
    key: string;
    text: string;
}
export interface IMultiSelectProp {
    label: string; //Label
    selectedItemIds?: string[]; //Ids of Selected Items
    onload: () => Promise<IItemProp[]>; //On load function to items for drop down 
    onPropChange: (targetProperty: string, oldValue: any, newValue: any) => void; // On Property Change function
    properties: any; //Web Part properties
    key?: string;  //unique key
    disabled: boolean;
}
 
export interface IMultiSelectPropInternal extends IPropertyPaneCustomFieldProps {
    targetProperty: string;
    label: string;
    selectedItemIds?: string[];
    onload: () => Promise<IItemProp[]>;
    onPropChange: (targetPropery: string, oldValue: any, newValue: any) => void;
    onRender: (elem: HTMLElement) => void;
    onDispose: (elem: HTMLElement) => void;
    properties: any;
    selectedKey: string;
    disabled: boolean;
}
 
export class MultiSelectBuilder implements IPropertyPaneField<IMultiSelectPropInternal>{
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public properties: IMultiSelectPropInternal;
 
    private label: string;
    private selectedItemIds: string[] = [];
    private onLoad: () => Promise<IItemProp[]>;
    private onPropChange: (targetPropery: string, oldValue: any, newValue: any) => void;
    private key: string;
    private cumstomProperties: any;
    private disabled: boolean;

    constructor(targetProperty: string, prop: IMultiSelectPropInternal) {
        this.targetProperty = targetProperty;
        this.properties = prop;
        this.label = prop.label;
        this.selectedItemIds = prop.selectedItemIds;
        this.cumstomProperties = prop.properties;
        this.onLoad = prop.onload.bind(this);
        this.onPropChange = prop.onPropChange.bind(this);
        this.properties.onRender = this.render.bind(this);
        this.properties.onDispose = this.dispose.bind(this);
        this.key = prop.key;
        this.disabled = prop.disabled;
    }
    
    public render(elem: HTMLElement, ctx?: any, changeCallback?: (targetProperty: string, oldValue: any, newValue: any) => void): void {
        let element: React.ReactElement<IMultiSelectHostProp> = React.createElement(MultiSelectHost, {
            targetProperty: this.targetProperty,
            label: this.label,
            properties: this.cumstomProperties,
            selectedItemIds: this.selectedItemIds,
            onDispose: null,
            onRender: null,
            onPropChange: changeCallback,
            onload: this.onLoad.bind(this),
            selectedKey: this.key,
            key: this.key,
            disabled: this.disabled,
            stateKey: new Date().toString()
        });
        ReactDom.render(element, elem);
    }
 
    private dispose(elem: HTMLElement): void {
    }
}
 
export function PropertyPaneMultiSelect(targetProperty: string, properties: IMultiSelectProp): IPropertyPaneField<IMultiSelectPropInternal> {
    const multiSelectProp: IMultiSelectPropInternal = {
        targetProperty: targetProperty,
        label: properties.label,
        properties: properties.properties,
        selectedItemIds: properties.selectedItemIds,
        onDispose: null,
        onRender: null,
        onPropChange: properties.onPropChange.bind(this),
        onload: properties.onload.bind(this),
        selectedKey: properties.key,
        key: properties.key,
        disabled: properties.disabled
    };
    return new MultiSelectBuilder(targetProperty, multiSelectProp);
}