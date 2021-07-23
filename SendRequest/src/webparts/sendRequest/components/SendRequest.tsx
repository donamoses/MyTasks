import * as React from 'react';
import styles from './SendRequest.module.scss';
import { ISendRequestProps } from './ISendRequestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Checkbox, DatePicker, DefaultButton, Dropdown, FontWeights, getTheme, IconButton, IDropdownOption, IDropdownStyles, IIconProps, Label, mergeStyleSets, PrimaryButton, TextField } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

import { sp, Web, View, ContentType } from "@pnp/sp/presets/all";

export interface ISendRequestState {
  currentuser: any;
  verifierId: any;
  Reviewer: any;
  approver: any;
  LinkToDoc: any;
  dcc: any;
  hideproject: boolean;
}
export default class SendRequest extends React.Component<ISendRequestProps, ISendRequestState, {}> {
  public constructor(props: ISendRequestProps) {
    super(props);
    this.state = {
      currentuser: "",
      verifierId: "",
      Reviewer: "",
      approver: "",
      LinkToDoc: "",
      dcc: "",
      hideproject: true
    };
  }
  public async componentDidMount() {

    await this.User();
    console.log(this.props.project);
    if (this.props.project) {
      this.setState({ hideproject: false });
    }
  }
  public async User() {
    let user = await sp.web.currentUser();
    this.setState({
      currentuser: user.Title,
    });
  }
  public _getVerifier = (items: any[]) => {

    console.log(items);
    let getSelectedUsers = [];

    for (let item in items) {
      getSelectedUsers.push(items[item].id);
    }
    this.setState({ verifierId: getSelectedUsers[0] });
    console.log(getSelectedUsers);
  }
  public render(): React.ReactElement<ISendRequestProps> {
    const controlClass = mergeStyleSets({
      control: {
        // margin    : '0 0 15px 0',
        maxWidth: '450px',
      },
    });
    const level: IDropdownOption[] = [

      { key: 'DIC', text: 'DIC' },
      { key: 'IDC', text: 'IDC' },
      { key: 'IFR', text: 'IFR' },
      { key: 'IFC', text: 'IFC' },
      { key: 'ABT', text: 'ABT' },
      { key: 'VOID', text: 'VOID' },
];
const BUOptions:IDropdownOption[]=[
  {key:'BU1',text:'BU1'}
];
const CategoryOptions:IDropdownOption[]=[
  {key:'BU1',text:'BU1'}
];
    const dropdownStyles: Partial<IDropdownStyles> = {
      dropdown: { width: 180 },
    };
    return (
      <div className={styles.sendRequest} >
         <div style={{ marginLeft: "auto",marginRight:"auto",width:"50rem" }}>
         <div className={styles.title}> Review and approval request form</div>
        
         
        <div >
          
          <Label >Document :  <a href={this.state.LinkToDoc}>NOT/SHML/INT-PRC/AM-00009 Migration Policy.docx</a></Label>
          <table>
            <tr>
              <td><Label >Orginator : SUNIL JOHN </Label></td>
              <td><Label >Requester : SUBHA RAVEENDRAN</Label></td>
              <td><Label >Revision : 0 </Label></td>
            </tr>
          </table>
          <table>
            <tr hidden={this.state.hideproject}>
              <td>
                <Dropdown id="RevisionLevel"
                  placeholder="Select an option"
                  label="Approval Level"
                  options={level}

                // styles={dropdownStyles}
                // selectedKey={this.state.selectedmin}
                // onChanged={(option) => this.min(option)}
                />
              </td>
              <td>
                <PeoplePicker
                  context={this.props.context}
                  titleText="DCC"
                  personSelectionLimit={1}
                  groupName={""} // Leave this blank in case you want to filter from all users    
                  showtooltip={true}

                  disabled={false}
                  ensureUser={true}
                  // selectedItems={this._getVerifier}
                  defaultSelectedUsers={[this.state.dcc]}
                  showHiddenInUI={false}
                  // isRequired={true}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                />
              </td>
            </tr>
          </table>
          <table>
            <tr><td>
                 <Dropdown id="t3" label="Business Unit"
                required={true}
                    placeholder="Select an option"
                    options={BUOptions}
                    // onChanged={this._drpdwnDepCateg} 
                    /></td>
                    
               <td> <Dropdown id="t2" required={true}label="Category"
                    placeholder="Select an option"
                    options={CategoryOptions}
                    // onChanged={this._drpdwnDocCateg}
                     />
</td></tr>
            <tr>
              <td>
                <PeoplePicker
                  context={this.props.context}
                  titleText="Reviewer(s)"
                  personSelectionLimit={8}
                  groupName={""} // Leave this blank in case you want to filter from all users    
                  showtooltip={true}
                  disabled={false}
                  ensureUser={true}
                  // selectedItems={this._getVerifier}
                  defaultSelectedUsers={[this.state.Reviewer]}
                  showHiddenInUI={false}
                  // isRequired={true}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                />
              </td>
              <td>
                <PeoplePicker
                  context={this.props.context}
                  titleText="Approver"
                  personSelectionLimit={1}
                  groupName={""} // Leave this blank in case you want to filter from all users    
                  showtooltip={true}
                  required={true}
                  disabled={false}
                  ensureUser={true}
                  // selectedItems={this._getApprover}
                  defaultSelectedUsers={[this.state.approver]}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} />
              </td>
            </tr>
          </table>
          <table>
            <tr>
              
              <td>
                <DatePicker label="Due Date:" id="DueDate" style={{ width: '300px' }}
                  //formatDate={(date) => moment(date).format('DD/MM/YYYY')}
                  isRequired={true}
                  // value={this.state.ExpireDate}
                  minDate={new Date()}
                  // className={controlClass.control}
                  // onSelectDate={this._onDatePickerChange}
                  placeholder="Due Date"
                />
              </td>
            </tr>
          </table>
          <table>

            <tr><td> <TextField label="Comments" id="Comments" multiline autoAdjustHeight /></td></tr>
            <tr><td hidden={this.state.hideproject}><Checkbox label="Approve in same revision ? " boxSide="end" /></td></tr>
          </table>
<div style={{padding:"0 0 0 38rem"}} >
  <Label style={{ color: "red",fontStyle:"italic",fontSize:"12px" }}>* fields are mandatory </Label>
  </div>
          <br />
          <DefaultButton id="b1" style={{ marginTop: '20px', float: "right", borderRadius: "10px", border: "1px solid gray" }}>Cancel</DefaultButton >
          <DefaultButton id="b2" style={{ marginTop: '20px', float: "right", marginRight: "10px", borderRadius: "10px", border: "1px solid gray" }}>Submit</DefaultButton >
          {/* <DefaultButton id="b2" style={{ marginTop: '20px', float: "right", marginRight: "10px", borderRadius: "10px", border: "1px solid gray" }}>Save</DefaultButton > */}
          <br />
          <br />

        </div>
        </div>  
      </div>
    );
  }
}
