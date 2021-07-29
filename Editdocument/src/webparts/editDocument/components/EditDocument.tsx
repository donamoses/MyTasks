import * as React from 'react';
import styles from './EditDocument.module.scss';
import { IEditDocumentProps } from './IEditDocumentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Checkbox, DatePicker, DefaultButton, DialogFooter, Dropdown, ITooltipHostStyles, Label, Pivot, PivotItem, TextField, TooltipHost } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
const calloutProps = { gapSpace: 0 };
const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { IFrameDialog } from "@pnp/spfx-controls-react/lib/IFrameDialog";
import { sp } from '@pnp/sp';
import "@pnp/sp/sites";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { VerticalTimeline, VerticalTimelineElement } from 'react-vertical-timeline-component';
import 'react-vertical-timeline-component/style.min.css';
import WorkIcon from '@material-ui/icons/Work';
import SchoolIcon from '@material-ui/icons/School';
import StarIcon from '@material-ui/icons/Star';
export interface IEditDocumentState {

  docs: any[];
  hidecreate:boolean;
  hideedit:boolean;
  siteurl:any;
  hideproject:boolean;
}
export default class EditDocument extends React.Component<IEditDocumentProps,IEditDocumentState, {}> {
  constructor(props: IEditDocumentProps) {
    super(props);
    this.state = {
       
       docs: [],
       hidecreate:false,
       hideedit:true,
       siteurl:"",
       hideproject:true

    };

}
public async componentDidMount() {
 console.log(this.props.project);
  
  const rootwebData = await sp.site.rootWeb();
  console.log(rootwebData);
  var webValue = rootwebData.ResourcePath.DecodedUrl;
  //alert(webValue);
 
  this.setState({
    siteurl: webValue
  });
  console.log(this.props.createdocument);
if (this.props.createdocument) {
  this.setState({ hidecreate: true,hideedit:false });
}
if(this.props.project){
  this.setState({hideproject:false});
}
this.getVersionHistory();
}
public getVersionHistory(){
  
}
  public render(): React.ReactElement<IEditDocumentProps> {
    return (
      <div className={ styles.editDocument }>
         <div>
                    <Pivot aria-label="Large Link Size Pivot Example">
                        <PivotItem headerText="Document Info">
                        <div style={{ marginLeft: "auto",marginRight:"auto",width:"30rem" }}>
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
                        <IFrameDialog
                        url={this.state.siteurl + "/_layouts/15/Versions.aspx?list=%7Bda53146b-3f5c-4321-926e-c3c2adbff323%7D&ID=1&IsDlg=0"}
                        title="Version History"
                        hidden={false}
                        // onDismiss={this.onCancel}
                        modalProps={{
                            isBlocking: true,
                            styles: { main: { maxWidth: "700px !important", width: "600px !important", height: "800px !important" } }
                        }}
                        dialogContentProps={{
                            type: DialogType.close,
                            showCloseButton: true
                        }}
                        width={'800px'}
                        height={'500px'}
                    />

                        </PivotItem>
                        <PivotItem headerText="Revision History">
                        <div style={{ width: "80%" }}>
                        
        <div> 
<VerticalTimeline>
  <VerticalTimelineElement
    className="vertical-timeline-element--work"
    contentStyle={{ background: 'rgb(233 157 127)', color: 'rgb(19 18 18)' }}
    contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
    date="24 Jul 2021"
    iconStyle={{ background: 'rgb(0, 120, 212)', color: '#fff' }}
    icon={<WorkIcon />}
  >
    <h3 style={{ fontSize: "20px", color: "rgb(220,20,60)" }}>Published</h3>
    <h4 className="vertical-timeline-element-subtitle">NOT/SHML/INT-PRC/AM-00009 Migration Policy.docx</h4>
    <p style={{fontSize:'12px'}}>
        Verified By : Subha Raveendran
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Approved By : Sunil John
        </div>
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Version : 1
        </div>
        <br></br>
      </p>
  </VerticalTimelineElement>
  <VerticalTimelineElement
    className="vertical-timeline-element--work"
    contentStyle={{ background: 'rgb(233 205 126)', color: 'rgb(19 18 18)' }}
    contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
    date="23 Jul 2021 "
    iconStyle={{ background: 'rgb(33, 150, 243)', color: '#fff' }}
    icon={<WorkIcon />}
  >
    <h3 style={{ fontSize: "20px", color: "rgb(220,20,60)" }}>Approved</h3>
    <h4 className="vertical-timeline-element-subtitle">NOT/SHML/INT-PRC/AM-00009 Migration Policy.docx</h4>
    <p style={{fontSize:'12px'}}>
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Approved By :  Sunil John
        </div>
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Version : 0
        </div>
        <br></br>
      </p>
  </VerticalTimelineElement>
  <VerticalTimelineElement
    className="vertical-timeline-element--work"
    contentStyle={{ background: 'rgb(213 202 231)', color: 'rgb(19 18 18)' }}
      contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
      date="22 Jul 2021"
      iconStyle={{ background: 'rgb(0, 120, 212)', color: '#fff' }}
    icon={<WorkIcon />}
  >
    <h3 style={{ fontSize: "20px", color: "rgb(220,20,60)" }}>Verified</h3>
    <h4 className="vertical-timeline-element-subtitle">NOT/SHML/INT-PRC/AM-00009 Migration Policy.docx</h4>
    <p style={{fontSize:'12px'}}>
        Requestor: Subha Raveendran
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Verified By : Subha Raveendran
        </div>
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Revision : 0
        </div>
        <br></br>
      </p>
  </VerticalTimelineElement>
  <VerticalTimelineElement
    className="vertical-timeline-element--work"
    contentStyle={{ background: 'rgb(185 237 137)', color: 'rgb(19 18 18)' }}
      contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
      date="21 Jul 2021"
      iconStyle={{ background: 'rgb(0, 120, 212)', color: '#fff' }}
    icon={<WorkIcon />}
  >
    <h3 style={{ fontSize: "20px", color: "rgb(220,20,60)" }}>WorkFlow Started</h3>
    <h4 className="vertical-timeline-element-subtitle">NOT/SHML/INT-PRC/AM-00009 Migration Policy.docx</h4>
    <p style={{fontSize:'12px'}}>
        <div> Requestor : Subha Raveendran</div>
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Verifier : Subha Raveendran  </div>
            <div style={{ margin: "0px 0px 0px 22px" }}>
              Approver : Sunil John</div>

       
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Revision : 0
        </div>
        <br></br>
      </p>
  </VerticalTimelineElement>
  <VerticalTimelineElement
    className="vertical-timeline-element--education"
    contentStyle={{ background: 'rgb(155 216 235 / 65%);', color: 'rgb(19 18 18)' }}
    contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
    date="20 Jul 2021"
    iconStyle={{ background: 'rgb(0, 120, 212)', color: '#fff' }}
    icon={<SchoolIcon />}
  >
    <h3 style={{ fontSize: "20px", color: "rgb(220,20,60)" }}> Document Created </h3>
    <h4 className="vertical-timeline-element-subtitle">NOT/SHML/INT-PRC/AM-00009 Migration Policy.docx</h4>
    <p style={{fontSize:'12px'}}>
<div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
  Originator :Sunil John
</div>
<div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>

</div>
<br></br>
</p>
  </VerticalTimelineElement>
  
 
  <VerticalTimelineElement
    iconStyle={{ background: 'rgb(16, 204, 82)', color: '#fff' }}
    icon={<StarIcon />}
  />

</VerticalTimeline>
        </div>
                          </div>
                        </PivotItem>
                        
                        {/* <PivotItem headerText="Transmittal History" >
                        <div style={{ width: "80%" }}>
                        
                        <div> 
                  <VerticalTimeline>
                  <VerticalTimelineElement
                    className="vertical-timeline-element--work"
                    contentStyle={{ background: 'rgb(233 157 127)', color: 'rgb(19 18 18)' }}
                    contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
                    date="24 Jul 2021"
                    iconStyle={{ background: 'rgb(0, 120, 212)', color: '#fff' }}
                    icon={<WorkIcon />}
                  >
                    <h3 style={{ fontSize: "20px", color: "rgb(220,20,60)" }}>Published</h3>
                    <h4 className="vertical-timeline-element-subtitle">NOT/SHML/INT-PRC/AM-00009 Migration Policy.docx</h4>
                    <p style={{fontSize:'12px'}}>
                        Verified By : Subha Raveendran
                        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                          Approved By : Sunil John
                        </div>
                        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                          Version : 1
                        </div>
                        <br></br>
                      </p>
                  </VerticalTimelineElement>
                  <VerticalTimelineElement
                    className="vertical-timeline-element--work"
                    contentStyle={{ background: 'rgb(233 205 126)', color: 'rgb(19 18 18)' }}
                    contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
                    date="23 Jul 2021 "
                    iconStyle={{ background: 'rgb(33, 150, 243)', color: '#fff' }}
                    icon={<WorkIcon />}
                  >
                    <h3 style={{ fontSize: "20px", color: "rgb(220,20,60)" }}>Approved</h3>
                    <h4 className="vertical-timeline-element-subtitle">NOT/SHML/INT-PRC/AM-00009 Migration Policy.docx</h4>
                    <p style={{fontSize:'12px'}}>
                        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                          Approved By :  Sunil John
                        </div>
                        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                          Version : 0
                        </div>
                        <br></br>
                      </p>
                  </VerticalTimelineElement>
                  <VerticalTimelineElement
                    className="vertical-timeline-element--work"
                    contentStyle={{ background: 'rgb(213 202 231)', color: 'rgb(19 18 18)' }}
                      contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
                      date="22 Jul 2021"
                      iconStyle={{ background: 'rgb(0, 120, 212)', color: '#fff' }}
                    icon={<WorkIcon />}
                  >
                    <h3 style={{ fontSize: "20px", color: "rgb(220,20,60)" }}>Verified</h3>
                    <h4 className="vertical-timeline-element-subtitle">NOT/SHML/INT-PRC/AM-00009 Migration Policy.docx</h4>
                    <p style={{fontSize:'12px'}}>
                        Requestor: Subha Raveendran
                        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                          Verified By : Subha Raveendran
                        </div>
                        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                          Revision : 0
                        </div>
                        <br></br>
                      </p>
                  </VerticalTimelineElement>
                  <VerticalTimelineElement
                    className="vertical-timeline-element--work"
                    contentStyle={{ background: 'rgb(185 237 137)', color: 'rgb(19 18 18)' }}
                      contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
                      date="21 Jul 2021"
                      iconStyle={{ background: 'rgb(0, 120, 212)', color: '#fff' }}
                    icon={<WorkIcon />}
                  >
                    <h3 style={{ fontSize: "20px", color: "rgb(220,20,60)" }}>WorkFlow Started</h3>
                    <h4 className="vertical-timeline-element-subtitle">NOT/SHML/INT-PRC/AM-00009 Migration Policy.docx</h4>
                    <p style={{fontSize:'12px'}}>
                        <div> Requestor : Subha Raveendran</div>
                        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                          Verifier : Subha Raveendran  </div>
                            <div style={{ margin: "0px 0px 0px 22px" }}>
                              Approver : Sunil John</div>
                
                       
                        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                          Revision : 0
                        </div>
                        <br></br>
                      </p>
                  </VerticalTimelineElement>
                  <VerticalTimelineElement
                    className="vertical-timeline-element--education"
                    contentStyle={{ background: 'rgb(155 216 235 / 65%);', color: 'rgb(19 18 18)' }}
                    contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
                    date="20 Jul 2021"
                    iconStyle={{ background: 'rgb(0, 120, 212)', color: '#fff' }}
                    icon={<SchoolIcon />}
                  >
                    <h3 style={{ fontSize: "20px", color: "rgb(220,20,60)" }}> Document Created </h3>
                    <h4 className="vertical-timeline-element-subtitle">NOT/SHML/INT-PRC/AM-00009 Migration Policy.docx</h4>
                    <p style={{fontSize:'12px'}}>
                
                <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                  Originator :Sunil John
                </div>
                <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                
                </div>
                <br></br>
                </p>
                  </VerticalTimelineElement>
                  <VerticalTimelineElement
                    iconStyle={{ background: 'rgb(16, 204, 82)', color: '#fff' }}
                    icon={<StarIcon />}
                  />
                </VerticalTimeline>
                        </div>
                                          </div>
                        </PivotItem> */}
                        
                    </Pivot>
                </div>
      </div>
    );
  }
}
