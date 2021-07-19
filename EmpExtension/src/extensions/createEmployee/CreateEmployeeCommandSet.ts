import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import {
  Panel,
  PanelType
} from 'office-ui-fabric-react';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import CreateEmployee from "../createEmployee/components/CreateEmployee";
import EditEmployeeForm from "../createEmployee/components/EditEmployeeForm";
import * as strings from 'CreateEmployeeCommandSetStrings';
import { ICreateEmployeeProps } from './components/ICreateEmployeeProps';
import { assign } from '@fluentui/react';
import { sp } from "@pnp/sp";
import * as $ from 'jquery';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICreateEmployeeCommandSetProperties {

  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
  sourceRelativeUrl: string;
}

const LOG_SOURCE: string = 'CreateEmployeeCommandSet';

export default class CreateEmployeeCommandSet extends BaseListViewCommandSet<ICreateEmployeeCommandSetProperties> {
  private panelPlaceHolder: HTMLDivElement = null;
  @override
  public onInit(): Promise<void> {
    this.properties.sourceRelativeUrl = "/sites/CCSHR/Lists/Employees";
    var Libraryurl = this.context.pageContext.list.serverRelativeUrl;
    Log.info(LOG_SOURCE, 'Initialized CreateEmployeeCommandSet');
    sp.setup({
      spfxContext: this.context
    });
    this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));//panel componet appending 
    if (Libraryurl == this.properties.sourceRelativeUrl) {
      // code to hide button
      setInterval(() => {
        $("button[name='New']").hide();
        $("button[name='Copy link']").hide();
        $("button[name='Share']").hide();
        $("button[name='Edit in grid view']").hide();
        $("button[name='Export to Excel']").hide();
        $("button[name='Power Apps']").hide();
        $("button[name='Automate']").hide();
        // $("button[aria-label='More']").hide();
        $("button[name='Comment']").hide();
        $("button[name='Edit']").hide();
        $("button[name='Alert me']").hide();
        $("button[name='Manage my alerts']").hide();
        $("button[name='Select items']").hide();
        $("button[name='Export']").hide();
        $("button[name='Integrate']").hide();
      }, 1);
    }

    return Promise.resolve();


  }
  public _showPanel() {
    this._renderPanelComponent({
      isOpen: true,
      listId: this.context.pageContext.list.id.toString(),//guid
      onClose: this._dismissPanel
    });
  }
  private _dismissPanel = () => {
    this._renderPanelComponent({ isOpen: false });
  }

  public _renderPanelComponent(props: any) {
    const element: React.ReactElement<ICreateEmployeeProps> = React.createElement(CreateEmployee, assign({
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

  public _showEditPanel() {
    this._renderEditPanelComponent({
      isOpen: true,




      listId: this.context.pageContext.list.id.toString(),
      onClose: this._dismissEditPanel
      //onClose: this._dismissPanel
    });

  }
  public _renderEditPanelComponent(props: any) {
    const element: React.ReactElement<ICreateEmployeeProps> = React.createElement(EditEmployeeForm, assign({
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
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    this.properties.sourceRelativeUrl = "/sites/CCSHR/Lists/Employees";
    var Libraryurl = this.context.pageContext.list.serverRelativeUrl;
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    const compareTwoCommand: Command = this.tryGetCommand('COMMAND_2');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = (event.selectedRows.length === 1 && (Libraryurl == this.properties.sourceRelativeUrl));
    }
    if (compareTwoCommand) {
      // compareTwoCommand.visible = (Libraryurl == this.properties.sourceRelativeUrl) && (this.visitorflag != 1);
      if (Libraryurl == this.properties.sourceRelativeUrl) {
        compareTwoCommand.visible = true;
      }
      else {
        compareTwoCommand.visible = false;
      }
    }
    if ((Libraryurl == this.properties.sourceRelativeUrl)) {
      setTimeout(() => {
        $("button[name='New']").hide();
        $("button[name='Copy link']").hide();
        $("button[name='Share']").hide();
        $("button[name='Edit in grid view']").hide();
        $("button[name='Export to Excel']").hide();
        $("button[name='Power Apps']").hide();
        $("button[name='Automate']").hide();
        // $("button[aria-label='More']").hide();
        $("button[name='Comment']").hide();
        $("button[name='Edit']").hide();
        $("button[name='Alert me']").hide();
        $("button[name='Manage my alerts']").hide();
        $("button[name='Select items']").hide();
        $("button[name='Export']").hide();
        $("button[name='Integrate']").hide();

      }, 2);
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    let Location;
    let Department;
    let Designation;
    let Nationality;
    let DOB;
    let DOJ;
    let Qualification;
    let ReportingOfficer;
    switch (event.itemId) {
      case 'COMMAND_1':
        if (event.selectedRows.length >= 1) {
          event.selectedRows.forEach(async (row: RowAccessor, index: number) => {
            let selectedItem = event.selectedRows[0];
            try {

              Location = row.getValueByName('Location')[0].lookupId;

            }
            catch {

              Location = null;

            }
            try {

              Qualification = row.getValueByName('Qualification')[0].lookupId;

            }
            catch {

              Qualification = null;

            }
            try {

              Department = selectedItem.getValueByName('Department')[0].lookupId;

            }
            catch {

              Department = null;

            }
            try {

              Designation = row.getValueByName('Designation')[0].lookupId;

            }
            catch {

              Designation = null;

            }
            try {

              Nationality = row.getValueByName('Nationality')[0].lookupId;

            }
            catch {

              Nationality = null;


            }
            try {

              ReportingOfficer = row.getValueByName('ReportingOfficer')[0].lookupId;

            }
            catch {

              ReportingOfficer = null;

            }
            if ((row.getValueByName('DOJ')) == "") {
              DOJ = null;

            }
            else {

              DOJ = new Date(row.getValueByName('DOJ'));

            }
            if ((row.getValueByName('DOB')) == "") {
              DOB = null;

            }
            else {

              DOB = new Date(row.getValueByName('DOB'));

            }
            let Usernameid;
            let Usernametitle;
            if (row.getValueByName('UserName') !== "") {
              row.getValueByName('UserName').forEach(elementss => {
                console.log(elementss);

                Usernametitle = elementss.title;
                Usernameid = elementss.id;
              });
            }
            const element: React.ReactElement<ICreateEmployeeProps> = React.createElement(EditEmployeeForm, assign({
              itemId: row.getValueByName('ID'),
              Firstname: row.getValueByName('FirstName'),
              Nationality: Nationality,
              AlternativeEmail: row.getValueByName('AlternativeEmail'),
              AlternativePhone: row.getValueByName('AlternativePhone'),
              WorkPhone: row.getValueByName('WorkPhone'),
              Country: row.getValueByName('Country'),
              City: row.getValueByName('City'),
              PIN: row.getValueByName('PIN'),
              PassportDetails: row.getValueByName('PassportDetails'),
              IDProofDetails: row.getValueByName('IDProofDetails'),
              DrivingLicense: row.getValueByName('DrivingLicense'),
              Address: row.getValueByName('Address'),
              NoticePeriod: row.getValueByName('NoticePeriod_x0028_inmonths_x002'),
              DateOfbirth: DOB,
              DateofJoining: DOJ,
              state: row.getValueByName('state'),
              Lastname: row.getValueByName('LastName'),
              Gender: row.getValueByName('Gender'),
              Department: Department,
              Designation: Designation,
              Location: Location,
              PermanentAddress: row.getValueByName('PermanentAddress'),
              TemporaryAddress: row.getValueByName('TemporaryAddress'),
              MobileNo: row.getValueByName('MobileNo'),
              EmailId: row.getValueByName('EmailId'),
              State: row.getValueByName('State'),
              Usernameid: Usernameid,
              Qualification: Qualification,
              ReportingOfficer: ReportingOfficer,
            }));
            ReactDom.render(element, this.panelPlaceHolder);

            this._showEditPanel();
          });
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
