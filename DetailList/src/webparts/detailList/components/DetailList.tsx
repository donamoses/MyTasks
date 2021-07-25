import * as React from 'react';
import styles from './DetailList.module.scss';
import { IDetailListProps } from './IDetailListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Item, Items, ITermSetInfo, sp } from "@pnp/sp/presets/all";
import { DetailsList, DetailsListLayoutMode, IColumn, IObjectWithKey, ISelection, Selection } from 'office-ui-fabric-react';
import { useRef } from 'react';
export interface IDetailListState {
  docRepositoryItems: any[];
  selectionDetails: string;
  items: any[];
}
export default class DetailList extends React.Component<IDetailListProps, IDetailListState, {}> {
  private _columns: IColumn[];
  private _selection: Selection;
  constructor(props: IDetailListProps) {
    super(props);
    this.state = {
      docRepositoryItems: [],
      selectionDetails: "",
      items: [],
    };
    this._columns = [
      { key: 'column1', name: 'Document Name', fieldName: 'DocumentName', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column2', name: 'Work Flow Status', fieldName: 'WFStatus', minWidth: 100, maxWidth: 200, isResizable: true },
    ];

  }
  public async componentDidMount() {

    this.loadDocProfile();

  }
  private loadDocProfile = async () => {
    //getting list DocProfile u
    sp.web.getList("/sites/DMS/Lists/DocumentProfile").items.get().then(docProfileItems => {

      this.setState({
        docRepositoryItems: docProfileItems,
        items: docProfileItems,
      });
      console.log(this.state.docRepositoryItems);
    });

  }
  private _onItemInvoked = (item) => {
    alert(item.ID);

  };


  public render(): React.ReactElement<IDetailListProps> {
    return (
      <div className={styles.detailList}>

        <DetailsList
          items={this.state.docRepositoryItems}
          columns={this._columns}
          layoutMode={DetailsListLayoutMode.justified}
          selection={this._selection}
          selectionPreservedOnEmptyClick={true}
          ariaLabelForSelectionColumn="Toggle selection"
          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          checkButtonAriaLabel="select row"

          onItemInvoked={item => this._onItemInvoked(item)}

        />

      </div>
    );
  }
}
