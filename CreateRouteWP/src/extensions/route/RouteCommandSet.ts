import { override } from '@microsoft/decorators';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Log } from '@microsoft/sp-core-library';
import { sp } from "@pnp/sp";
import { assign } from '@uifabric/utilities';
import * as $ from 'jquery';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import CreateRoute from "../components/CreateRoute";
import EditRoute from "../components/EditRoute";
import { IRouteProps } from "../components/IRouteProps";

import * as strings from 'RouteCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IRouteCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'RouteCommandSet';

export default class RouteCommandSet extends BaseListViewCommandSet<IRouteCommandSetProperties> {
  private panelPlaceHolder: HTMLDivElement = null;

  @override
  public onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });
    this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));
    Log.info(LOG_SOURCE, 'Initialized RouteCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }
  public _showPanel() {
    this._renderPanelComponent({
      isOpen: true,
      // paneltype: "Medium",
      //currentTitle,
      //itemId,
      listId: this.context.pageContext.list.id.toString(),
      onClose: this._dismissPanel
      //onClose: this._dismissPanel
    });
  }
  public _showEditPanel() {
    this._renderEditPanelComponent({
      isOpen: true,

      // paneltype: "",
      //currentTitle,
      //itemId,


      listId: this.context.pageContext.list.id.toString(),
      onClose: this._dismissEditPanel
      //onClose: this._dismissPanel
    });

  }
  private _dismissPanel = () => {

    this._renderPanelComponent({ isOpen: false });
  }
  public _renderPanelComponent(props: any) {
    const element: React.ReactElement<IRouteProps> = React.createElement(CreateRoute, assign({
      onClose: null,
      paneltype: "",
      //onClose: null,
      // currentTitle: null,
      // itemId: null,
      isOpen: false,
      context: this.context
      //  listId: null
    }, props));


    ReactDom.render(element, this.panelPlaceHolder);
  }
  public _renderEditPanelComponent(props: any) {
    const element: React.ReactElement<IRouteProps> = React.createElement(EditRoute, assign({
      onClose: null,
      paneltype: "",
      isOpen: false,
      context: this.context
    }, props));
    ReactDom.render(element, this.panelPlaceHolder);

  }
  public _dismissEditPanel = () => {
    this._renderEditPanelComponent({ isOpen: false });
  }


  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    let PlannedDatefromlist;
    let Districtfromlist;
    let Dealernamefromlist;
    let contactnumberfromlist;
    let locationfromlist;
    let assigntofromlist;
    let remarksfromlist;
    let PlannedVisitTimefromlist;
    let dealerarray = [];
    let assigntoarray = [];
    switch (event.itemId) {
      case 'COMMAND_1':
        if (event.selectedRows.length > 0) {
          event.selectedRows.forEach(async (row: RowAccessor, index: number) => {
            if ((row.getValueByName('PlannedDateFormatted')) == "") {
              PlannedDatefromlist = null;

            }
            else {

              PlannedDatefromlist = new Date(row.getValueByName('PlannedDateFormatted'))

            }
            try {

              Districtfromlist = row.getValueByName('District')[0].lookupId;

            }
            catch {

              Districtfromlist = null;

            }
            try {

              Dealernamefromlist = row.getValueByName('DealerName')[0].lookupId;

            }
            catch {

              Dealernamefromlist = null;

            }
            if ((row.getValueByName('ContactNumber')) == null) {
              contactnumberfromlist = null;

            }
            else {

              contactnumberfromlist = row.getValueByName('ContactNumber');

            }
            try {

              locationfromlist = row.getValueByName('Location')[0].lookupId;

            }
            catch {

              locationfromlist = null;

            }
            try {

              assigntofromlist = row.getValueByName('AssignTo')[0].lookupId;

            }
            catch {

              assigntofromlist = null;

            }
            if ((row.getValueByName('Title')) == '') {
              PlannedVisitTimefromlist = '';

            }
            else {

              PlannedVisitTimefromlist = row.getValueByName('Title');

            }
            if ((row.getValueByName('Remarks')) == "") {
              remarksfromlist = "";

            }
            else {

              remarksfromlist = row.getValueByName('Remarks').replace(/(<([^>]+)>)/gi, "");

            }
            const dealeritems: any[] = await sp.web.lists.getByTitle("Dealer List").items.select("Title,ID").filter(" DistrictId eq " + Districtfromlist).get();
            console.log("dealer" + dealeritems);
            for (let i = 0; i < dealeritems.length; i++) {

              let data = {
                key: dealeritems[i].Id,
                text: dealeritems[i].Title
              };

              dealerarray.push(data);
            }
            const salesuseritems: any[] = await sp.web.lists.getByTitle("Users").items.select("Title,ID").filter(" DistrictId eq " + Districtfromlist).get();
            console.log("salesusers" + salesuseritems);
            for (let i = 0; i < salesuseritems.length; i++) {

              let data = {
                key: salesuseritems[i].Id,
                text: salesuseritems[i].Title
              };

              assigntoarray.push(data);
            }

            const element: React.ReactElement<IRouteProps> = React.createElement(EditRoute, assign({
              itemidprops: row.getValueByName('ID'),
              PlannedDateprops: PlannedDatefromlist,
              Districtprops: Districtfromlist,
              DealerNameprops: Dealernamefromlist,
              ContactNumberprops: contactnumberfromlist,
              Locationprops: locationfromlist,
              AssignToprops: assigntofromlist,
              PlannedVisitTimeprops: PlannedVisitTimefromlist,
              Remarksprops: remarksfromlist,
              dealeroptionsprops: dealerarray,
              assigntooptionprops: assigntoarray
            }));
            ReactDom.render(element, this.panelPlaceHolder);
            this._showEditPanel();
          })
        }

        break;
      case 'COMMAND_2':
        this._showPanel();
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
