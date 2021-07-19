import * as React from 'react';
import styles from './Outlook2SharePoint.module.scss';
import * as strings from 'Outlook2SharePointWebPartStrings';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import GraphController from '../../../controller/GraphController';
import { IOutlook2SharePointProps } from './IOutlook2SharePointProps';
import { IOutlook2SharePointState } from './IOutlook2SharePointState';
import Select from 'react-select';
import { sp } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-groups/web";
import { IHubSiteWebData } from "@pnp/sp/hubsites";
import { DetailsList, DetailsListLayoutMode, Selection, IColumn, SelectionMode, IDetailsListProps, IDetailsRowStyles, DetailsRow } from 'office-ui-fabric-react/lib/DetailsList';
import "@pnp/sp/hubsites/web";
import "@pnp/sp/hubsites";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
import { ActionButton, Breadcrumb, IBreadcrumbItem, Image, ImageFit, Link, PrimaryButton, Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { IFolder } from '../../../model/IFolder';
import { getTheme } from 'office-ui-fabric-react/lib/Styling';
import { _Items } from '@pnp/sp/items/types';
import { _Folder } from '@pnp/sp/folders/types';
import Utilities from '../../../controller/Utilities';
const theme = getTheme();
const colourStyles = {
  control: styles2 => ({ ...styles2, width: '270px' }),
  option: (styles2) => {
    return {
      ...styles2
    };
  },
};
const GetImgUrlByFileExtension = (extension: string) => {
  // cuurently in SPFx with React I didn't find different way of getting the image
  // feel free to improve this
  let imgRoot: string = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001//assets/item-types/20/";
  let imgType = "genericfile.png";
  imgType = extension + ".png";

  switch (extension) {
    case "jpg":
    case "jpeg":
    case "jfif":
    case "gif":
    case "png":
      imgType = "photo.png";
      break;
    case "SP.Folder":
      imgType = "folder.svg";
      break;

  }
  return imgRoot + imgType;
};
var hubsiteArray = [];
var listArrays = [];
var siteCollectionArray = [];
export default class Outlook2SharePoint extends React.Component<IOutlook2SharePointProps, IOutlook2SharePointState> {
  private breadItems: IBreadcrumbItem[] = [];
  private graphController: GraphController;
  private selectRef = null;
  private _selection: Selection;
  private saveMetadata = true; // For simplicity reasons and as I am not convinced with the current "Property handling" of Office Add-In we configure 'hard-coded'
  private relativeUrlForFetchingList: string;
  private _columns: IColumn[];
  constructor(props) {
    super(props);
    // this._selection = new Selection({
    //   onSelectionChanged: () => { this._onItemEdit(); }
    // });
    this.state = {
      graphController: null,
      mailMetadata: null,
      showError: false,
      showSuccess: false,
      showOneDrive: false,
      showTeams: false,
      showGroups: false,
      successMessage: '',
      errorMessage: '',
      selectedHubOptions: null,
      hubsiteSelected: { value: "", label: "" },
      listSelected: { value: "", label: "" },
      siteSelected: { value: "", label: "" },
      folderStructureItems: [],
      currentRootFolder: "",
      key: 0,
      initialFocusedIndex: 0,
      folders: [],
      grandParentFolder: null,
      parentFolder: null,
      selectedGroupName: '',
      showSpinner: false,
      selectedFolder: null,
      hideLoader: "none"
    };
    this.graphController = new GraphController(this.saveMetadata);
    this.graphController.init(this.props.msGraphClientFactory)
      .then((controllerReady) => {
        if (controllerReady) {
          this.graphClientReady();
        }
      });
    this.relativeUrlForFetchingList = `${this.props.webServerRelativeUrl}/Lists/`;
    this._columns = [
      { key: 'odata.type', name: '', fieldName: 'odata.type', minWidth: 20, maxWidth: 50, isResizable: false },
      { key: 'Name', name: 'Folder', fieldName: 'Name', minWidth: 100, maxWidth: 200, isResizable: true }
    ];

  }
  public async componentDidMount() {
    await this._getAllHubSites();
  }
  private handleSelectSiteCollection = async (selectedSiteCollections) => {
    let govDocList: [];
    this.breadItems = [];
    this.selectRef.value = selectedSiteCollections;
    //Load document library
    let siteWeb = Web(selectedSiteCollections.value);
    this.getGroups(selectedSiteCollections.label).then(async loadGroups => {
      govDocList = await siteWeb.getFolderByServerRelativeUrl('Shared Documents').expand("Folders/ListItemAllFields").get();
      if (this.state.currentRootFolder == "") {
        this.breadItems.push({ text: govDocList["Name"], key: govDocList["ServerRelativeUrl"], onClick: this._onBreadcrumbItemClicked });
      }
      let array = govDocList["Folders"].filter(excludeForm).sort((a, b) => {
        if (b.Name > a.Name)
          return -1;
        if (b.Name < a.Name)
          return 1;
        return 0;
      });
      this.selectRef.value = null;
      return this.setState({
        siteSelected: { value: selectedSiteCollections.value, label: selectedSiteCollections.label },
        folderStructureItems: array,
        currentRootFolder: govDocList["ServerRelativeUrl"]
      });
    });
  }
  private _onBreadcrumbItemClicked = async (ev: React.MouseEvent<HTMLElement>, item: IBreadcrumbItem) => {
    let siteWeb = Web(this.state.siteSelected.value);
    let govDocList = [];
    govDocList = await siteWeb.getFolderByServerRelativeUrl(item.key).expand("Folders/ListItemAllFields").get();
    let requiredIndex = 0;
    this.breadItems.map((bItems, index) => {
      if (bItems.key == item.key) {
        requiredIndex = index;
      }
    });
    this.breadItems.length = Number(requiredIndex) + 1;
    this.setState({
      folderStructureItems: govDocList["Folders"].filter(excludeForm),
      currentRootFolder: govDocList["ServerRelativeUrl"]
    });
  }
  private handleHubSites = async (selectedHubSites) => {
    this.selectRef.value = selectedHubSites;
    let web = Web(selectedHubSites.value);
    listArrays = [];
    await web.lists.get().then(listName => {
      console.log(listName);
        for (let i in listName) {
          listArrays.push({
            value: listName[i].Title,
            label: listName[i].Title,
          });
        }
      });
    this.selectRef.value = null;
    return this.setState({
      hubsiteSelected: { value: selectedHubSites.value, label: selectedHubSites.label }
    });
  }
  private handleList = async (selectedList) => {
    this.selectRef.value = selectedList;
    siteCollectionArray = [];
    let govDocList: [];
    this.breadItems = [];
    // let subWeb = Web(this.state.hubsiteSelected.value);
    // await subWeb.lists
    //   .getByTitle(selectedList.value).items.
    //   getAll()
    //   .then((listItems) => {
    //     for (let k in listItems) {
    //       if (listItems[k].Link != null) {
    //         siteCollectionArray.push({
    //           value: listItems[k].Link.Url != null ? listItems[k].Link.Url.trim() : "",
    //           label: listItems[k].Link.Url != null ? listItems[k].Title : "Link is empty",
    //         });
    //       }
    //     }
    //   });
    // this.selectRef.value = null;
    // return this.setState({
    //   listSelected: { value: selectedList.value, label: selectedList.label }
    // });
    
    let siteWeb = Web(this.state.hubsiteSelected.value);
    this.getGroups(selectedList.label).then(async loadGroups => {
      govDocList = await siteWeb.getFolderByServerRelativeUrl(selectedList.label).expand("Folders/ListItemAllFields").get();
      if (this.state.currentRootFolder == "") {
        this.breadItems.push({ text: govDocList["Name"], key: govDocList["ServerRelativeUrl"], onClick: this._onBreadcrumbItemClicked });
      }
      let array = govDocList["Folders"].filter(excludeForm).sort((a, b) => {
        if (b.Name > a.Name)
          return -1;
        if (b.Name < a.Name)
          return 1;
        return 0;
      });
      this.selectRef.value = null;
      return this.setState({
        listSelected: { value: selectedList.value, label: selectedList.label },
        folderStructureItems: array,
        currentRootFolder: govDocList["ServerRelativeUrl"]
      });
    });
  }
  /**
*  Get hub site. used to bind values in dropdown
*/
  private async _getAllHubSites() {
    let adminWeb = Web("https://ccsdev01.sharepoint.com/sites/Portal");
    hubsiteArray = [];
    adminWeb.hubSiteData().then(results => {
      for (var i = 0; i < results.navigation.length; i++) {
        hubsiteArray.push({
          value: results.navigation[i].Url,
          label: results.navigation[i].Title,
        });
      }
    });
  }
  /**
   * This function first retrieves all OneDrive root folders from user
   */
  private graphClientReady = () => {
    this.setState((prevState: IOutlook2SharePointState, props: IOutlook2SharePointProps) => {
      return {
        graphController: this.graphController
      };
    });
    if (this.saveMetadata) {
      this.getMetadata();
    }
  }


  private closeMessage = () => {
    this.setState((prevState: IOutlook2SharePointState, props: IOutlook2SharePointProps) => {
      return {
        showSuccess: false,
        showError: false
      };
    });
  }

  private getMetadata() {
    this.state.graphController.retrieveMailMetadata(this.props.mail.id)
      .then((response) => {
        if (response !== null) {
          this.setState((prevState: IOutlook2SharePointState, props: IOutlook2SharePointProps) => {
            return {
              mailMetadata: response
            };
          });
        }
      });
  }
  //Setting up Bread crumb navigation
  private _navigate = async (items: string) => {
    let siteWeb = Web(this.state.siteSelected.value);
    let govDocList = [];
    govDocList = await siteWeb.getFolderByServerRelativeUrl(items['ServerRelativeUrl']).expand("Folders/ListItemAllFields").get();
    console.log(govDocList);
    this.breadItems.push({ text: items["Name"], key: items["ServerRelativeUrl"], onClick: this._onBreadcrumbItemClicked });
    let array = govDocList["Folders"].filter(excludeForm).sort((a, b) => {
      if (b.Name > a.Name)
        return -1;
      if (b.Name < a.Name)
        return 1;
      return 0;
    });
    let filterSubFolder = this.state.parentFolder.filter(item => {
      return item.name == items["Name"];
    });
    if (filterSubFolder != undefined && filterSubFolder != null) {
      await this.state.graphController.getSubFolder(filterSubFolder).then(subFolders => {
        this.setState({
          folders: subFolders,
          grandParentFolder: null,
          selectedFolder: filterSubFolder,
          parentFolder: subFolders,
          hideLoader: "none",
          folderStructureItems: array,
          currentRootFolder: items['ServerRelativeUrl'],
          initialFocusedIndex: 0
        });
      });
    }
  }
  private _renderItemColumn = (item: any, index: number, column: IColumn) => {
    let fieldContent = item[column.key];
    switch (column.key) {
      case 'odata.type':
        return < Image src={GetImgUrlByFileExtension(fieldContent)} width={32} height={32} imageFit={ImageFit.center}></Image>;
      case 'Name':
        return <Link onClick={() => this._navigate(item)} style={{ margin: '7px 0 0 0' }}>{item[column.key]}</Link>;
      default:
        return <span>{fieldContent}</span>;
    }
  }
  private getGroups = async (groupName: string) => {
    // await this.state.graphController.getGroupID(groupName).then(groupID => {
      this.state.graphController.getGroupDrives(groupName).then(async (folders) => {
        console.log(folders);
        await this.state.graphController.getSubFolder(folders).then(subFolders => {
          this.setState({
            folders: subFolders,
            grandParentFolder: null,
            parentFolder: subFolders
          });
        });
      });
    // });
  }
  private async _onItemEdit() {
    this.setState({
      hideLoader: "block"
    });
    const selectionCount = this._selection.getSelection();
    let filterSubFolder = this.state.parentFolder.filter(item => {
      return item.name == selectionCount[0]["Name"];
    });
    if (filterSubFolder != undefined && filterSubFolder != null) {
      await this.state.graphController.getSubFolder(filterSubFolder).then(subFolders => {
        this.setState({
          folders: subFolders,
          grandParentFolder: null,
          selectedFolder: filterSubFolder,
          parentFolder: subFolders,
          hideLoader: "none"
        });
      });
    }
  }
  private saveMailTo = () => {
    console.log("Yes");
    this.setState({
      hideLoader: "block",
    });
    console.log(this.state.selectedFolder[0].driveID);
    this.state.graphController.retrieveMimeMail(this.state.selectedFolder[0].driveID, this.state.selectedFolder[0].id, this.props.mail, this.saveMailCallback)
      .then((response) => {
        let siteName = this.state.siteSelected.value;
        this.state.graphController.retrieveMailContent(this.state.selectedFolder[0].driveID, this.state.selectedFolder[0].id, this.props.mail, this.saveMailCallback)
          .then(mailBody => {
            if (response.length < (4 * 1024 * 1024))      // If Mail size bigger 4MB use resumable upload
            {
              this.state.graphController.saveNormalMail(this.state.selectedFolder[0].driveID, this.state.selectedFolder[0].id, response, Utilities.createMailFileName(this.props.mail.subject), this.saveMailCallback).then(webURL => {
                mailBody = mailBody["bodyPreview"] + "<a href='" + webURL + "' target='blank'>Link</a>";
                this.state.graphController.GetSiteID(siteName.split('/Sites/')[1]).then(siteDetails => {
                  this.state.graphController.savetoLog(siteDetails.id, this.props.mail, mailBody).then(savedSuccessfully => {
                    this.setState({
                      hideLoader: "none",
                      showSuccess: true,
                    });
                  });
                });
              });
            }
            else {
              this.state.graphController.saveBigMail(this.state.selectedFolder[0].driveID, this.state.selectedFolder[0].id, response, Utilities.createMailFileName(this.props.mail.subject), this.saveMailCallback).then(saveBigMailResponse => {
                mailBody = mailBody["bodyPreview"] + "<a href='" + saveBigMailResponse + "' target='blank'>Link</a>";
                this.state.graphController.GetSiteID(siteName.split('/Sites/')[1]).then(siteDetails => {
                  this.state.graphController.savetoLog(siteDetails.id, this.props.mail, mailBody).then(savedSuccessfully => {
                    this.setState({
                      hideLoader: "none",
                      showSuccess: true,
                    });
                  });
                });
              });
            }
          });
      });
  }
  private saveMailCallback = (message: string) => {
    if (message.indexOf('Success') > -1) {
      this.props.successCallback(strings.SuccessMessage);
      this.setState({
        hideLoader: "none",
      });
    }
    else {
      this.props.errorCallback(strings.ErrorMessage);
      this.setState({
        hideLoader: "none",
      });
    }
  }
  public render(): React.ReactElement<IOutlook2SharePointProps> {
    return (
      <div className={styles.outlook2SharePoint} style={{ position: 'fixed' }}>
        {this.state.mailMetadata !== null &&
          <div className={styles.metadata}>
            <div><Icon iconName="InfoSolid" /> {strings.SaveInfo}</div>
            <div className={styles.subMetadata}>{strings.To} <a href={this.state.mailMetadata.saveUrl}>{this.state.mailMetadata.saveDisplayName}</a></div>
            <div className={styles.subMetadata}>{strings.On} <span>{this.state.mailMetadata.savedDate.toLocaleDateString()}</span></div>
          </div>}
        <div style={{ marginLeft: '8px' }}>
        <label></label>
          <label>Select Hub</label>
          <Select
            ref={(scope) => { this.selectRef = scope; }}
            label={"Hub"}
            isSearchable={true}
            //defaultValue={this.state.hubsiteSelected}
            onChange={(selectedOption) =>
              this.handleHubSites(selectedOption)
            }
            value={this.state.hubsiteSelected}
            placeholder={"Select hub"}
            options={hubsiteArray}
            styles={colourStyles}
            required={true}
          />
          <label>Select List</label>
          <Select
            ref={(scope) => { this.selectRef = scope; }}
            label={"List"}
            isSearchable={true}
            onChange={(selectedOption) =>
              this.handleList(selectedOption)
            }
            value={this.state.listSelected}
            placeholder={"Select site"}
            options={listArrays}
            styles={colourStyles}
            required={true}
          />
          {/* <label>Select Site</label>
          <Select
            ref={(scope) => { this.selectRef = scope; }}
            label={"Site"}
            isSearchable={true}
            onChange={(selectedOption) =>
              this.handleSelectSiteCollection(selectedOption)
            }
            value={this.state.siteSelected}
            placeholder={"Select site"}
            options={siteCollectionArray}
            styles={colourStyles}
            required={true}
          /> */}
        </div>
        {this.state.showSuccess && <div>
          <MessageBar
            messageBarType={MessageBarType.success}
            isMultiline={false}
            onDismiss={this.closeMessage}
            dismissButtonAriaLabel="Close"
            truncated={true}
            overflowButtonAriaLabel="See more"
          >
            Successfully copied to Site
          </MessageBar>
        </div>}
        {this.state.showError && <div>
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
            onDismiss={this.closeMessage}
            dismissButtonAriaLabel="Close"
            truncated={true}
            overflowButtonAriaLabel="See more"
          >
            Something went wrong
          </MessageBar>
        </div>}
        <div>
          <Breadcrumb
            items={this.breadItems}
            maxDisplayedItems={10}
          />
          <DetailsList
            items={this.state.folderStructureItems}
            columns={this._columns}
            onItemInvoked={this._navigate}
            onRenderItemColumn={this._renderItemColumn}
            // selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            compact={true}
            initialFocusedIndex={this.state.initialFocusedIndex}
            selectionMode={SelectionMode.none}
          />
        </div>
        <div>
          <PrimaryButton
            style={{ width: "30px", margin: '11px 0px', float: 'right' }}
            text="Copy"
            onClick={this.saveMailTo}
            allowDisabledFocus={true}
          />
          <Spinner style={{ display: this.state.hideLoader }} size={SpinnerSize.large} />
        </div>
      </div>
    );
  }
}
function excludeForm(form) {
  return form.Name != "Forms";
}