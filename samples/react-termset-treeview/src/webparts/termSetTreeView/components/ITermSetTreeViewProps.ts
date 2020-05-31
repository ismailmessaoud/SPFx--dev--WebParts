import { ISPTermSets } from 'sp-client-custom-fields/lib/PropertyFieldTermSetPicker';
import { ITreeItem } from "@pnp/spfx-controls-react/lib/TreeView";

export interface ITermSetTreeViewProps {
  description: string;
  termSet:ISPTermSets;
  treeItems:ITreeItem[];
  defaultExpanded:boolean;
  selectionMode:number;
  selectChildrenIfParentSelected:boolean;
  showCheckboxes:boolean;
  treeItemActionsDisplayMode:number;
}
