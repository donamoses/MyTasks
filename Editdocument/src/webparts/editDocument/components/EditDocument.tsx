import * as React from 'react';
import styles from './EditDocument.module.scss';
import { IEditDocumentProps } from './IEditDocumentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Checkbox, DatePicker, DefaultButton, DialogFooter, Dropdown, ITooltipHostStyles, Label, Pivot, PivotItem, TextField, TooltipHost } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
const calloutProps = { gapSpace: 0 };
const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
export interface IEditDocumentState {

  docs: any[];
  hidecreate:boolean;
  hideedit:boolean
}
export default class EditDocument extends React.Component<IEditDocumentProps,IEditDocumentState, {}> {
  constructor(props: IEditDocumentProps) {
    super(props);
    this.state = {
       
        docs: [],
       hidecreate:false,
       hideedit:true,

    };

}
public async componentDidMount() {
  console.log(this.props.createdocument);
if (this.props.createdocument) {
  this.setState({ hidecreate: true,hideedit:false });
}
}
  public render(): React.ReactElement<IEditDocumentProps> {
    return (
      <div className={ styles.editDocument }>
         <div>
                    <Pivot aria-label="Large Link Size Pivot Example">
                        <PivotItem headerText="Document Info">
                        <div style={{ marginLeft: "auto",marginRight:"auto",width:"50rem" }}>
                        <div style={{fontSize:"18px",fontWeight:"bold",textAlign:"center"}}> Edit Document</div>
                        < TextField required id="t1"
                          label="Title"
                          // onKeyUp={this._titleValidation}
                          // onChange={this._titleChange}
                          value="" >
                        </TextField>
                        <PeoplePicker
                    context={this.props.context}
                    titleText="Originator"
                    personSelectionLimit={1}
                    groupName={""} // Leave this blank in case you want to filter from all users    
                    showtooltip={true}
                    required={false}
                    disabled={false}
                    ensureUser={true}
                    // onChange={this._getDocResponsible}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000} />
                     <PeoplePicker
                    context={this.props.context}
                    titleText="Reviewer(s)"
                    personSelectionLimit={8}
                    groupName={""} // Leave this blank in case you want to filter from all users    
                    showtooltip={true}
                    required={false}
                    disabled={false}
                    ensureUser={true}
                    // onChange={this._Verifier}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    // defaultSelectedUsers={[this.state.setverifier]}
                    resolveDelay={1000} />
                <PeoplePicker
                    context={this.props.context}
                    titleText="Approver"
                    personSelectionLimit={3}
                    groupName={""} // Leave this blank in case you want to filter from all users    
                    showtooltip={true}
                    required={false}
                    disabled={false}
                    ensureUser={true}
                    // onChange={this._Approver}
                    showHiddenInUI={false}
                    // defaultSelectedUsers={[this.state.setapprover]}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000} />
                     <DatePicker label="Expiry Date"
                   style={{ width: '200px' }}
                    // value=""
                    // onSelectDate={this._onExpDatePickerChange}
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                />
                <div hidden={this.state.hidecreate}>
                 <Label >Select a Template:</Label>  <Dropdown id="t7"
                    placeholder="Select an option"

                    options={this.state.docs} 
                    // onChanged={this.templatechange}
                />
                <Label >Upload Document:</Label> <input  type="file" id="myfile" ></input>
                </div>
                <table>
                        <tr>
                            <td hidden={this.state.hidecreate} >
                                <TooltipHost
                                content="Check if the template or attachment is added"
                                //id={tooltipId}
                                calloutProps={calloutProps}
                                styles={hostStyles}>
                                    <Checkbox label="Create Document ? " boxSide="end"  />
                                </TooltipHost>
                            </td>
                            <td hidden={this.state.hideedit}>
                                <TooltipHost
                                content="Check if the template or attachment is added"
                                //id={tooltipId}
                                calloutProps={calloutProps}
                                styles={hostStyles}>
                                    <Checkbox label="Create Document ? " boxSide="end" defaultChecked />
                                </TooltipHost>
                            </td>
                            <td style={{width:"2rem"}}></td>
                            <td> 
                                <TooltipHost
                                content="The document to published library without sending it for review/approval"
                                //id={tooltipId}
                                calloutProps={calloutProps}
                                styles={hostStyles}>
                                    <Checkbox label="Direct Publish ? " boxSide="end" />
                                </TooltipHost>
                            </td>
                        </tr>
                    </table>
                    <DialogFooter>
                    {/* <PrimaryButton text="Save" onClick={this._onCreateDocument} />
                    <PrimaryButton text="Cancel" onClick={this._onCancel} /> */}
                    <DefaultButton id="b1" style={{ marginTop: '20px', float: "right", borderRadius: "10px", border: "1px solid gray" }}>Cancel</DefaultButton >
                    <DefaultButton id="b2" style={{ marginTop: '20px', float: "right", marginRight: "10px", borderRadius: "10px", border: "1px solid gray" }}>Submit</DefaultButton >

                </DialogFooter>
                        </div>
                           
                        </PivotItem>
                        <PivotItem headerText="Version History">
                           

                        </PivotItem>
                        <PivotItem headerText="Revision History">
                           
                        </PivotItem>
                        <PivotItem headerText="Transmittal History">
                           
                        </PivotItem>
                    </Pivot>
                </div>
      </div>
    );
  }
}
