import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TermSetTreeViewWebPartStrings';
import TermSetTreeView from './components/TermSetTreeView';
import { ITermSetTreeViewProps } from './components/ITermSetTreeViewProps';
import { PropertyFieldTermSetPicker } from 'sp-client-custom-fields/lib/PropertyFieldTermSetPicker';
import { ISPTermSets } from 'sp-client-custom-fields/lib/PropertyFieldTermSetPicker';
import { ITreeItem } from "@pnp/spfx-controls-react/lib/TreeView";
import SPTaxonomyService from '../../services/SPTaxonomyService';
import { ITermData, ITerm  } from "@pnp/sp-taxonomy";

export interface ITermSetTreeViewWebPartProps {
  description: string;
  termSet: ISPTermSets;
  treeItems:ITreeItem[];
  defaultExpanded:boolean;
  selectionMode:number;
  selectChildrenIfParentSelected:boolean;
  showCheckboxes:boolean;
  treeItemActionsDisplayMode:number;
}

export default class TermSetTreeViewWebPart extends BaseClientSideWebPart <ITermSetTreeViewWebPartProps> {
  private spTaxonomyService: SPTaxonomyService = new SPTaxonomyService();
  public render(): void {
    const element: React.ReactElement<ITermSetTreeViewProps> = React.createElement(
      TermSetTreeView,
      {
        description: this.properties.description,
        termSet: this.properties.termSet,
        treeItems: this.properties.treeItems,
        defaultExpanded: this.properties.defaultExpanded,
        selectionMode: this.properties.selectionMode,
        selectChildrenIfParentSelected: this.properties.selectChildrenIfParentSelected,
        showCheckboxes: this.properties.showCheckboxes,
        treeItemActionsDisplayMode: this.properties.treeItemActionsDisplayMode,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private fetchTermSetTags(): Promise<ITreeItem[]>{
    const wp = this;
    if(this.properties.termSet){
      var options: Array <ITreeItem> = new Array <ITreeItem> ();
      return(this.spTaxonomyService.getTermSetTags(this.context.pageContext.site.absoluteUrl,this.properties.termSet[0].TermStoreGuid,this.properties.termSet[0].Guid)).then((response) => {
        response.map((term: ITermData & ITerm) => {
          options.push({
            key: wp.spTaxonomyService.cleanGuid(term.Id),
            label: term.Name,
            data:term['Parent'] ? wp.spTaxonomyService.cleanGuid(term['Parent'].Id) :null
          });
        });
        return options;
      });
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath == 'termSet' && newValue) {
      this.fetchTermSetTags().then((response) => {
        this.properties.treeItems = this.spTaxonomyService.getTermSetTree(response,{idKey:'key',parentKey:'data'});
        this.render();
      });
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyFieldTermSetPicker('termSet', {
                  label: strings.TermSetFieldLabel,
                  panelTitle: strings.TermSetFieldPanel,
                  initialValues: this.properties.termSet,
                  allowMultipleSelections: false,
                  excludeSystemGroup: false,
                  displayOnlyTermSetsAvailableForTagging: false,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  render: this.render.bind(this),
                  disableReactivePropertyChanges: this.disableReactivePropertyChanges,
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'termSetPickerFieldId'
                }),
              ]
            },
            {
              groupName: strings.TreeViewGroupName,
              groupFields: [
                PropertyPaneToggle('defaultExpanded', {
                  label: strings.DefaultExpandedFieldLabel,
                  onText: strings.DefaultExpandedonText,
                  offText: strings.DefaultExpandedoffText
                }),
                PropertyPaneDropdown('selectionMode', {
                  label: strings.SelectionModeFieldLabel,
                  options: [
                    { key: '0', text: strings.SelectionModeSingle },
                    { key: '1', text: strings.SelectionModeMultiple },
                    { key: '2', text: strings.SelectionModeNone },
                  ]
                }),
                PropertyPaneToggle('selectChildrenIfParentSelected', {
                  label: strings.SelectChildrenFieldLabel,
                  onText: strings.SelectChildrenonText,
                  offText: strings.SelectChildrenoffText
                }),
                PropertyPaneCheckbox('showCheckboxes', {
                  text: strings.ShowCheckboxesFieldText
                }),
                PropertyPaneDropdown('treeItemActionsDisplayMode', {
                  label: "Display mode of the tree item actions",
                  options: [
                    { key: '1', text: strings.DisplayModeButtons },
                    { key: '2', text: strings.DisplayModeContextualMenu },
                  ]
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
