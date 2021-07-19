import * as React from 'react';
import styles from './SendRequest.module.scss';
import { ISendRequestProps } from './ISendRequestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Checkbox, DatePicker, DefaultButton, Dropdown, FontWeights, getTheme, IconButton, IDropdownOption, IDropdownStyles, IIconProps, Label, mergeStyleSets, PrimaryButton, TextField } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

import { sp, Web, View, ContentType } from "@pnp/sp/presets/all";
const theme = getTheme();
const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexFlow: 'column nowrap',
    alignItems: 'stretch',
    height: '700px'
  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      // borderTop: `10px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      padding: '12px 12px 14px 24px',

    },
  ],
  body: {
    width: '750px',
    flex: '4 4 auto',
    padding: '0 24px 24px 24px',
    overflowY: 'hidden',
    selectors: {
      p: { margin: '14px 0' },
      'p:first-child': { marginTop: 0 },
      'p:last-child': { marginBottom: 0 },
    },
  },
});
const cancelIcon: IIconProps = { iconName: 'Cancel' };
const iconButtonStyles = {
  root: {
    color: theme.palette.neutralPrimary,
    marginLeft: 'auto',
    marginTop: '4px',
    marginRight: '2px',
  },
  rootHovered: {
    color: theme.palette.neutralDark,
  },
};
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
    const dropdownStyles: Partial<IDropdownStyles> = {
      dropdown: { width: 180 },
    };
    return (
      <div className={styles.sendRequest}>
        <div className={contentStyles.header}>
          <span className={styles.title}>Review and approval request form of NOT/SHML/INT-PRC/AM-00009</span>
          <IconButton
            iconProps={cancelIcon}
            ariaLabel="Close popup modal"
            // onClick={this._closeModal}
            styles={iconButtonStyles}
          />
        </div>
        <div >
          <Label >Document :  <a href={this.state.LinkToDoc}>NOT/SHML/INT-PRC/AM-00009 Migration Policy</a></Label>
          <table>
            <tr>
              <td><Label >Orginator : {this.state.currentuser} </Label></td>
              <td><Label >Requester : {this.state.currentuser}</Label></td>
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
            </tr>
          </table>
          <table>
            <tr>
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
              <td>
                <DatePicker label="Due Date:" id="DueDate" style={{ width: '100%' }}
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
          <Label style={{ color:"red" }}>* fields are mandatory </Label>
          <br />
          <DefaultButton id="b1" style={{ marginTop: '20px', float: "right", borderRadius: "10px", border: "1px solid gray" }}>Cancel</DefaultButton >
          <DefaultButton id="b2" style={{ marginTop: '20px', float: "right", marginRight: "10px", borderRadius: "10px", border: "1px solid gray" }}>Submit</DefaultButton >
          <DefaultButton id="b2" style={{ marginTop: '20px', float: "right", marginRight: "10px", borderRadius: "10px", border: "1px solid gray" }}>Save</DefaultButton >
          <br />
          <br />

        </div>
      </div>
    );
  }
}
