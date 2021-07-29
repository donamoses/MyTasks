import * as React from 'react';
import styles from './CreateDocument.module.scss';
import { ICreateDocumentProps } from './ICreateDocumentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Checkbox, DatePicker, DefaultButton, DialogFooter, Dropdown, Icon, ITooltipHostStyles, Label, MessageBar, MessageBarType, PrimaryButton, TextField, TooltipHost } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from '@pnp/sp/presets/all';
import "@pnp/sp/sites";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { Web } from "@pnp/sp/webs";
import { IAttachmentFileInfo } from "@pnp/sp/attachments";
import * as moment from 'moment';
const calloutProps = { gapSpace: 0 };
const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
const MyIcon = () => <Icon iconName="TextDocumentSettings" />;
export interface ICreateDocumentState {
  saving: boolean;
  depOptions: any[];
  docCategoryOptions: any[];
  docs: any[];
  addUsers: any[];
  value: any;
  allItemss: any[];
  key: any;
  deptkey: any;
  verifierid: any;
  approverid: any;
  docRespId: any;
  tkey: any;
  setverifier: string;
  setapprover: string;
  // firstDayOfWeek?: DayOfWeek;
  expiredate: any;
  showProgress: boolean;
  progressLabel: string;
  progressDescription: string;
  progressPercent: any;
  userdept: any;
  depttext: any;
  siteurl: any;
  title: any;
  DocumentAdded: any;
}
export default class CreateDocument extends React.Component<ICreateDocumentProps,ICreateDocumentState,any> {
  constructor(props: ICreateDocumentProps) {
    super(props);
    this.state = {
        saving: false,
        depOptions: [],
        docCategoryOptions: [],
        docs: [],
        addUsers: [],
        showProgress: false,
        progressLabel: "File upload progress",
        progressDescription: "",
        progressPercent: 0,
        value: "",
        allItemss: [],
        key: "",
        deptkey: "",
        verifierid: "",
        approverid: "",
        docRespId: "",
        tkey: "",
        setverifier: "",
        setapprover: "",
        expiredate: "",
        userdept: "",
        depttext: "",
        siteurl: "",
        title: "",
        DocumentAdded: 'none',
    };

    this._drpdwnDocCateg = this._drpdwnDocCateg.bind(this);
    this._drpdwnDepCateg = this._drpdwnDepCateg.bind(this);
    this._onCancel = this._onCancel.bind(this);
    this.templatechange = this.templatechange.bind(this);
    this._titleValidation = this._titleValidation.bind(this);
}
public async componentDidMount() {
this.getData();
}
public getData=async ()=>{
  let uweb = Web(this.props.EmployeeUrl);
        let userdeptvalue;
        let username;
        username = await sp.web.currentUser.get();
        // const EmployeeListitems: any[] = await uweb.lists.getByTitle("Employees").items.select("Department/Title").expand("Department").filter(" UserNameId eq " + username.Id).get();
        const EmployeeListitems: any[] = await uweb.lists.getByTitle(this.props.EmployeelistName).items.select("Title", "UserName/Id", "UserName/Title", "UserName/EMail", "EmailId","Department/Title").expand("UserName","Department").filter(" UserName/EMail eq  '" + username.Email + "'").get();
        console.log(EmployeeListitems);
        console.log(EmployeeListitems[0].Department.Title);
        userdeptvalue = EmployeeListitems[0].Department.Title;
        this.setState({ userdept: userdeptvalue });
        // alert(this.state.userdept);
        const allItems: any[] = await sp.web.lists.getByTitle(this.props.DepartmentlistName).items.select("Title,ID,DepartmentName").getAll();
        let optionsArray = [];
       
        for (let i = 0; i < allItems.length; i++) {

            let data = {
                key: allItems[i].Id,
                text: allItems[i].DepartmentName
            };

            optionsArray.push(data);
        }
        this.setState({
          depOptions: optionsArray
        });
        let optionsArrays = [];
        const allItemss: any[] = await sp.web.lists.getByTitle(this.props.DocumentlistName).items.select("DocumentCategory,ID").getAll();
  
  for (let i = 0; i < allItemss.length; i++) {

      let data = {
          key: allItemss[i].Id,
          text: allItemss[i].DocumentCategory
      };

      optionsArrays.push(data);
  }
  this.setState({
      docCategoryOptions: optionsArrays
  });
  console.log(this.state.docCategoryOptions);
  //Select Template Dropdown
  let docarray = [];
  let value = this.props.TemplateCategory;

  const Items: any[] = await sp.web.lists.getByTitle(this.props.TemplatelistName).items.select("DocumentName").filter("substringof('" + value + "',DocumentName)").get();
        //console.log(Items);
  for (let i = 0; i < Items.length; i++) {

      let data = {
          key: Items[i].DocumentName,
          text: Items[i].DocumentName
      };

      docarray.push(data);
  }
  this.setState({
      docs: docarray
  });
  //alert(Items[0].DocumentName);


}
public _titleValidation = () => {

  let titlemsg = ((document.getElementById("t1") as HTMLInputElement).value);
  let titleformat = /^[A-Za-z0-9\s]*$/;
  if (titlemsg.match(titleformat)) {
      document.getElementById("msg").style.display = 'none';
  }
  else {
      document.getElementById("msg").style.display = 'inline';

  }

}
private _titleChange = (ev: React.FormEvent<HTMLInputElement>, Title?: string) => {
  this.setState({ title: Title || '' });
}
public _drpdwnDocCateg(option: { key: any; }) {
  //console.log(option.key);
  this.setState({ key: option.key });
}
public _drpdwnDepCateg = async (option) => {
  console.log(option.key);
  this.setState({
      deptkey: option.key,
      depttext: option.text
  });

  const items: any[] = await sp.web.lists.getByTitle(this.props.DepartmentlistName).items.select("Approver/Title", "Approver/Id", "Verifier/Title", "Verifier/Id").expand("Approver", "Verifier").filter("substringof('" + option.text + "',DepartmentName)").get();
  console.log(items);
  console.log(items[0].Approver.Title);
  this.setState({ setapprover: items[0].Approver.Title });
  this.setState({ setverifier: items[0].Verifier.Title });
  this.setState({ verifierid: items[0].Verifier.Id });
  this.setState({ approverid: items[0].Approver.Id });
}
public _getDocResponsible = (items: any[]) => {
  console.log(items);
  let getSelectedUsers = [];
  for (let item in items) {
      getSelectedUsers.push(items[item].id);
  }
  this.setState({ docRespId: getSelectedUsers[0] });
  console.log(getSelectedUsers);


}
public _Verifier = (items: any[]) => {

  console.log(items);
  let getSelectedUsers = [];

  for (let item in items) {
      getSelectedUsers.push(items[item].id);
  }
  this.setState({ verifierid: getSelectedUsers[0] });
  console.log(getSelectedUsers);
  //this.setState({ addUsers: getSelectedUsers });
  //console.log(this.state.addUsers);

  //console.log('Items:', items);

}
public _Approver = (items: any[]) => {

  console.log(items);
  let getSelectedUsers = [];

  for (let item in items) {
      getSelectedUsers.push(items[item].id);
  }
  this.setState({ approverid: getSelectedUsers[0] });
  console.log(getSelectedUsers);
  //this.setState({ addUsers: getSelectedUsers });
  //console.log(this.state.addUsers);

  //console.log('Items:', items);

}
private _onExpDatePickerChange = (date?: Date): void => {

  this.setState({ expiredate: date });

}
public templatechange(option: { key: any; }) {
  //console.log(option.key);
  this.setState({ tkey: option.key });
}
private fileUpload = () => {
  let fileInfos: IAttachmentFileInfo[] = [];
  let input = document.getElementById("myfile") as HTMLInputElement;
  var fileCount = input.files.length;
  console.log(fileCount);
  for (var i = 0; i < fileCount; i++) {
      var fileName = input.files[i].name;
      console.log(fileName);
      var file = input.files[i];
      var reader = new FileReader();
      reader.onload = ((fileN => {
          return (e) => {
              console.log(fileN.name);
              //Push the converted file into array
              fileInfos.push({
                  "name": file.name,
                  "content": e.target.result,
              });
              console.log(fileInfos);
          };
      }))(file);
      reader.readAsArrayBuffer(file);
  }
  //End of for loop
}
public _onCreateDocument = async () => {
  this.fileUpload();
  console.log(this.state.tkey);
  let deptstatus;
  if (this.state.depttext == this.state.userdept) {
      deptstatus = "Yes";
  }
  else {
      deptstatus = "No";
  }

  //let title = ((document.getElementById("t1") as HTMLInputElement).value);
  let keyword = ((document.getElementById("keyword") as HTMLInputElement).value);
  let category = ((document.getElementById("t2") as HTMLInputElement).innerText);
  let doccategory = category.replace(/[^a-zA-Z0-9]/g, '');
  let dept = ((document.getElementById("t3") as HTMLInputElement).innerText);
  let dp = dept.replace(/[^a-zA-Z ]/g, "");
  let datee = moment(this.state.expiredate, 'DD/MM/YYYY').format("DD MMM YYYY");
  console.log(datee);
  let siteUrl = this.state.siteurl;
  let web = Web(siteUrl);
  let file = document.getElementById("myfile") as HTMLInputElement;
  let filess = file.files[0];
  console.log(this.state.docRespId);
  console.log(datee);
  if (this.state.docRespId == "" && datee == "Invalid date") {
    sp.web.lists.getByTitle(this.props.ListName).items.add({
          Title: this.state.title,
          DocumentCategoryId: this.state.key,
          DepartmentId: this.state.deptkey,
          //DocumentResponsibleId: this.state.docrespid,
          VerifierId: this.state.verifierid,
          ApproverId: this.state.approverid,
          TempalateDocument: this.state.tkey,
          KeywordSearch: keyword,
          UserDepartment: deptstatus,
          // Expiredate: datee
      }).then(async i => {
          if (filess != undefined) {
              i.item.attachmentFiles.add(filess.name, filess);
          }
          this.setState({ saving: false });
         

          console.log(i);
      });
  }
  if (this.state.docRespId != "" && datee != "Invalid date") {
    sp.web.lists.getByTitle(this.props.ListName).items.add({

          Title: this.state.title,
          DocumentCategoryId: this.state.key,
          DepartmentId: this.state.deptkey,
          DocumentResponsibleId: this.state.docRespId,
          VerifierId: this.state.verifierid,
          ApproverId: this.state.approverid,
          TempalateDocument: this.state.tkey,
          KeywordSearch: keyword,
          ExpiryDate: datee,
          UserDepartment: deptstatus,
      }).then(i => {
          if (filess != undefined) {
              i.item.attachmentFiles.add(filess.name, filess);
              if (filess.size <= 10485760) {
                  i.item.attachmentFiles.add(filess.name, filess);
              }
              else {
                  i.item.attachmentFiles.add(filess.name, filess);
                  i.item.attachmentFiles.add(filess.name, filess);
              }
          }
          this.setState({ saving: false });
          
          console.log(i);
      });
  }
  if (this.state.docRespId != "" && datee == "Invalid date") {
    sp.web.lists.getByTitle(this.props.ListName).items.add({
          Title: this.state.title,
          DocumentCategoryId: this.state.key,
          DepartmentId: this.state.deptkey,
          DocumentResponsibleId: this.state.docRespId,
          VerifierId: this.state.verifierid,
          ApproverId: this.state.approverid,
          TempalateDocument: this.state.tkey,
          KeywordSearch: keyword,
          UserDepartment: deptstatus,
          // ExpiryDate: datee
      }).then(i => {
          if (filess != undefined) {
              i.item.attachmentFiles.add(filess.name, filess);
              if (filess.size <= 10485760) {
                  i.item.attachmentFiles.add(filess.name, filess);
              }
              else {
                  i.item.attachmentFiles.add(filess.name, filess);
              }
          }
          this.setState({ saving: false });
          
          console.log(i);
      });
  }
  if (this.state.docRespId == "" && datee != "Invalid date") {
    sp.web.lists.getByTitle(this.props.ListName).items.add({
          Title: this.state.title,
          DocumentCategoryId: this.state.key,
          DepartmentId: this.state.deptkey,
          // DocumentResponsibleId: this.state.docrespid,
          VerifierId: this.state.verifierid,
          ApproverId: this.state.approverid,
          TempalateDocument: this.state.tkey,
          ExpiryDate: datee,
          KeywordSearch: keyword,
          UserDepartment: deptstatus,
      }).then(i => {
          if (filess != undefined) {
              i.item.attachmentFiles.add(filess.name, filess);
              if (filess.size <= 10485760) {
                  i.item.attachmentFiles.add(filess.name, filess);
              }
              else {
                  i.item.attachmentFiles.add(filess.name, filess);
              }
          }
          this.setState({ saving: false });
          
          console.log(i);
      });
  }
    // alert("Document Created Successfully");
    this.setState({ DocumentAdded: '' });
            setTimeout(() => this.setState({ DocumentAdded: 'none' }), 1000);

    this._onCancel();
}
private _onCancel = () => {
  window.location.href = this.props.RedirectUrl;
}
public render(): React.ReactElement<ICreateDocumentProps> {

    return (
      <div className={styles.createDocument}>
           <div style={{ marginLeft: "auto",marginRight:"auto",width:"50rem" }}>
         <div className={styles.alignCenter}> Create Document</div>
         <div style={{ display: this.state.DocumentAdded }}>
                    <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>  Document Saved Successfully.</MessageBar>
                </div>
                
                
                < TextField required id="t1"
                label="Name"
                 onKeyUp={this._titleValidation}
                  onChange={this._titleChange}
                   value={this.state.title} ></TextField>
                <div id="msg"><Label style={{ color: "green" }}>Name can't contain any of the following characters: \ /:*?""|&#{ }%~"</Label></div>
                

                 <Dropdown id="t3" label="Business Unit"
                required={true}
                    placeholder="Select an option"
                    options={this.state.depOptions}
                    onChanged={this._drpdwnDepCateg} />
                    
                <Dropdown id="t2" required={true}label="Category"
                    placeholder="Select an option"
                    options={this.state.docCategoryOptions}
                    onChanged={this._drpdwnDocCateg} />
                <Dropdown id="t2" required={true}label="Sub Category"
                    placeholder="Select an option"
                    options={this.state.docCategoryOptions}
                    onChanged={this._drpdwnDocCateg} />
                <PeoplePicker
                    context={this.props.context}
                    titleText="Originator"
                    personSelectionLimit={1}
                    groupName={""} // Leave this blank in case you want to filter from all users    
                    showtooltip={true}
                    required={false}
                    disabled={false}
                    ensureUser={true}
                    onChange={this._getDocResponsible}
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
                    onChange={this._Verifier}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    defaultSelectedUsers={[this.state.setverifier]}
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
                    onChange={this._Approver}
                    showHiddenInUI={false}
                    defaultSelectedUsers={[this.state.setapprover]}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000} />
                {/* <TooltipHost
                    content="Multiple Keywords should be ',' separated"
                    //id={tooltipId}
                    calloutProps={calloutProps}
                    styles={hostStyles}>
                    <Label >Keyword: </Label>< TextField id="keyword"   ></TextField>
                </TooltipHost> */}
                {/* </Tooltip> */}
                
               
                   
                <DatePicker label="Expiry Date"
                   style={{ width: '200px' }}
                    value={this.state.expiredate}
                    onSelectDate={this._onExpDatePickerChange}
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                />
               
               
               
                <Label >Select a Template:</Label>  <Dropdown id="t7"
                    placeholder="Select an option"

                    options={this.state.docs} onChanged={this.templatechange}
                />
                <Label >Upload Document:</Label> <input  type="file" id="myfile" ></input>
                <table>
                        <tr>
                            <td>
                                <TooltipHost
                                content="Check if the template or attachment is added"
                                //id={tooltipId}
                                calloutProps={calloutProps}
                                styles={hostStyles}>
                                    <Checkbox label="Create Document ? " boxSide="end" />
                                </TooltipHost>
                            </td><td style={{width:"2rem"}}></td>
                            <td> 
                                <TooltipHost
                                content="The document to published library without sending it for review/approval"
                                //id={tooltipId}
                                calloutProps={calloutProps}
                                styles={hostStyles}>
                                    <Checkbox label="Direct Publish ? " boxSide="end" />
                                </TooltipHost>
                            </td> <td style={{width:"2rem"}}></td>
                            <td> 
                                <TooltipHost
                                content="Is the document is Critical"
                                //id={tooltipId}
                                calloutProps={calloutProps}
                                styles={hostStyles}>
                                    <Checkbox label="Critical Document ? " boxSide="end" />
                                </TooltipHost>
                            </td>
                        </tr>
                    </table>
                    <div style={{padding:"0 0 0 38rem"}} >
  <Label style={{ color: "red",fontStyle:"italic",fontSize:"12px" }}>* fields are mandatory </Label>
  </div>
                <DialogFooter>
                    {/* <PrimaryButton text="Save" onClick={this._onCreateDocument} />
                    <PrimaryButton text="Cancel" onClick={this._onCancel} /> */}
                    <DefaultButton id="b1" style={{ marginTop: '20px', float: "right", borderRadius: "10px", border: "1px solid gray" }}>Cancel</DefaultButton >
                    <DefaultButton id="b2" style={{ marginTop: '20px', float: "right", marginRight: "10px", borderRadius: "10px", border: "1px solid gray" }}>Submit</DefaultButton >

                </DialogFooter>
                </div>    
      </div>
    );
  }
}
