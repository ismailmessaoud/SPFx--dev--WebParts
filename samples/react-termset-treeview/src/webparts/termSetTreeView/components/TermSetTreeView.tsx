import * as React from 'react';
import styles from './TermSetTreeView.module.scss';
import { ITermSetTreeViewProps } from './ITermSetTreeViewProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from 'TermSetTreeViewWebPartStrings';
import { TreeView, ITreeItem, TreeViewSelectionMode, TreeItemActionsDisplayMode } from "@pnp/spfx-controls-react/lib/TreeView";

export default class TermSetTreeView extends React.Component<ITermSetTreeViewProps, {}> {
  private treeItems:ITreeItem[] = [];
  
  public componentWillMount(): void {
    if(this.props.treeItems){
      this.treeItems = this.props.treeItems; 
    }
  }

  public componentWillUpdate (): void {
    if(this.props.treeItems){
      this.treeItems = this.props.treeItems; 
    }
  }
  
  private onExpandCollapseTree(item: ITreeItem, isExpanded: boolean) {
    console.log((isExpanded ? "Item expanded: " : "Item collapsed: ") + item);
  }

  private onItemSelected(items: ITreeItem[]) {
    console.log("Items selected: " + items.length);
    console.log(items);
  }

  public render(): React.ReactElement<ITermSetTreeViewProps> {
    return (
      <div >   
        <h1>{strings.WPTitle}</h1>  
        <TreeView
            items={this.treeItems}
            defaultExpanded={this.props.defaultExpanded}
            selectionMode={this.props.selectionMode}
            selectChildrenIfParentSelected={this.props.selectChildrenIfParentSelected}
            showCheckboxes={this.props.showCheckboxes}
            treeItemActionsDisplayMode={this.props.treeItemActionsDisplayMode}
            onExpandCollapse={this.onExpandCollapseTree}
            onSelect={this.onItemSelected} 
        />
        </div>
    );
  }
}
