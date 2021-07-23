import * as React from 'react';
import styles from './DocumentApproval.module.scss';
import { IDocumentApprovalProps } from './IDocumentApprovalProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton, Dropdown, IDropdownOption, Label, TextField } from 'office-ui-fabric-react';
export interface IDocumentApprovalState {
  requestor: any;
  LinkToDoc: any;
  requestorComments: any;
  dueDate: any;
  DCCComments: any;
  hideproject: boolean;
}
export default class DocumentApproval extends React.Component<IDocumentApprovalProps, IDocumentApprovalState, {}> {
  public constructor(props: IDocumentApprovalProps) {
    super(props);
    this.state = {

      requestor: "",
      LinkToDoc: "",
      requestorComments: "",
      dueDate: "",
      DCCComments: "",
      hideproject: true
    };
  }
  public async componentDidMount() {
      console.log(this.props.project);
    if (this.props.project) {
      this.setState({ hideproject: false });
    }
  }
  public render(): React.ReactElement<IDocumentApprovalProps> {
    const Status: IDropdownOption[] = [

      { key: 'Approved', text: 'Approved' },
      { key: 'Rejected', text: 'Rejected' },
     
    ];
    const PublishOption: IDropdownOption[] = [

      { key: 'PDF', text: 'PDF' },
      { key: 'Native', text: 'Native' },
     
    ];
    return (
      <div className={styles.documentApproval}>
         <div style={{ marginLeft: "auto",marginRight:"auto",width:"50rem" }}>
        <div className={styles.alignCenter}> Approval form</div>
        <div >
         
          <Label >Document :  <a href={this.state.LinkToDoc}>NOT/SHML/INT-PRC/AM-00009 Migration Policy.docx</a></Label>
          <table>
            <tr>
              <td><Label >Revision : 0 </Label></td>
              <td hidden={this.state.hideproject}><Label>Revision Level : ABT </Label></td>
            </tr>
            <tr>
              <td><Label >Orginator : SUNIL JOHN </Label></td>
              <td><Label >Due Date : 21 JUL 2021 </Label></td>
            </tr>
            <tr>
              <td><Label>Requestor : SUBHA RAVEENDRAN </Label></td>
              <td><Label >Requested Date : 21 JUL 2021 </Label></td></tr></table><table>
              <tr>
              <td><Label> Requestor Comment:</Label>Requested to review the document </td>
            </tr></table><table>
            <tr hidden={this.state.hideproject} >
              <td><Label>DCC : SUBHA RAVEENDRAN </Label></td>
              <td><Label >DCC Date : 21 JUL 2021 </Label></td></tr></table><table><tr hidden={this.state.hideproject}>
              <td><Label> DCC Comment:</Label>Requested to dcc level review the document</td>
            </tr></table><table>
            <tr  >
              <td><Label>Reviewer : SUBHA RAVEENDRAN </Label></td>
              <td><Label >Review Date : 21 JUL 2021 </Label></td></tr></table><table><tr>
              <td><Label> Review Comment:</Label>Requested to approve the document</td>
            </tr>
          </table>

        </div>
        <div style={{ marginTop: '30px' }}>
        <Dropdown 
          placeholder="Select Option" 
          label="Publish Option"
          style={{ marginBottom: '10px', backgroundColor: "white" }}
          options={PublishOption}
          // onChanged={this.ChangeId}
          // selectedKey={this.state.Status ? this.state.Status.key : undefined}
          required />
          <Dropdown 
          placeholder="Select Status" 
          label="Status"
          style={{ marginBottom: '10px', backgroundColor: "white" }}
          options={Status}
          // onChanged={this.ChangeId}
          // selectedKey={this.state.Status ? this.state.Status.key : undefined}
          required />

        <TextField label="Comments" id="Comments" multiline autoAdjustHeight />
        
        <div style={{padding:"0 0 0 38rem"}} >
  <Label style={{ color: "red",fontStyle:"italic",fontSize:"12px" }}>* fields are mandatory </Label>
  </div>
          <DefaultButton id="b1" style={{ marginTop: '20px', float: "right", borderRadius: "10px", border: "1px solid gray" }}>Cancel</DefaultButton >
          <DefaultButton id="b2" style={{ marginTop: '20px', float: "right", marginRight: "10px", borderRadius: "10px", border: "1px solid gray" }}>Submit</DefaultButton >
          <DefaultButton id="b2" style={{ marginTop: '20px', float: "right", marginRight: "10px", borderRadius: "10px", border: "1px solid gray" }}>Save</DefaultButton >
          <br />
        </div>
        </div>
      </div>
    );
  }
}
