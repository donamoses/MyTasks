import * as React from 'react';
import styles from './DocDetails.module.scss';
import { IDocDetailsProps } from './IDocDetailsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DetailsList, DetailsListLayoutMode, IColumn,Selection } from '@fluentui/react';
import { Item, sp, Web } from "@pnp/sp/presets/all";
export interface IDocDetailsState {
  docDetails: any[];
  items: IDetailsListBasicExampleItem[];
  selectionDetails: string;
}
export interface IDetailsListBasicExampleItem {
  key: number;
  name: string;
  value: number;
}
export default class DocDetails extends React.Component<IDocDetailsProps, IDocDetailsState, any> {
  private _columns: IColumn[];
  private _selection: Selection;
  private _allItems: IDetailsListBasicExampleItem[];
 
  docDetails: any;

  constructor(props: IDocDetailsProps) {
    super(props);
   
    this.state = ({
      docDetails: [],
      items:[],
      selectionDetails:""
    });
   
}
public async componentDidMount() {
  this.getDoc();
  }
  public getDoc=async ()=>{
    let uweb = Web(this.props.siteUrl);
    // const DocumentRepositoryItems: any[] = await sp.web.lists.getByTitle(this.props.listName).items.select("DocumentName,DocumentResponsible/Title,DocumentResponsible/ID,ID,Revision,Created,Author/ID,Author/Title,WFStatus,Approver/ID,Approver/Title,Verifier/ID,Verifier/Title").expand("DocumentResponsible,Author,Approver,Verifier").get();
    //     console.log(DocumentRepositoryItems);
          // Populate with items for demos.
     this._allItems = [];
     this._allItems = [];
    for (let i = 0; i < 200; i++) {
      this._allItems.push({
        key: i,
        name: 'Item ' + i,
        value: i,
      });
    }
 console.log(this._allItems);
     this._columns = [
       { key: 'column1', name: 'Name', fieldName: 'name', minWidth: 100, maxWidth: 200, isResizable: true },
       { key: 'column2', name: 'Value', fieldName: 'value', minWidth: 100, maxWidth: 200, isResizable: true },
     ];
    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
    });
    this.state = {
      docDetails: [],
      items: this._allItems,
      selectionDetails: this._getSelectionDetails(),
    };
  }
private _getSelectionDetails(): string {
  const selectionCount = this._selection.getSelectedCount();

  switch (selectionCount) {
    case 0:
      return 'No items selected';
    case 1:
      return '1 item selected: ' + (this._selection.getSelection()[0] as IDetailsListBasicExampleItem).name;
    default:
      return `${selectionCount} items selected`;
  }
}

private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
  this.setState({
    items: text ? this._allItems.filter(i => i.name.toLowerCase().indexOf(text) > -1) : this._allItems,
  });
};

private _onItemInvoked = (item: IDetailsListBasicExampleItem): void => {
  alert(`Item invoked: ${item.name}`);
};

  public render(): React.ReactElement<IDocDetailsProps> {
    const { items, selectionDetails } = this.state;
    return (
      <div className={styles.docDetails}>
          <DetailsList
            items={items}
            columns={this._columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="select row"
            onItemInvoked={this._onItemInvoked}
          />
      </div>
    );
  }
}
