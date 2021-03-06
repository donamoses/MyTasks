import { MSGraphClient, MSGraphClientFactory } from '@microsoft/sp-http';
import Utilities from './Utilities';
import { IFolder } from '../model/IFolder';
import { IMail } from '../model/IMail';
import { IMailMetadata } from '../model/IMailMetadata';

export default class GraphController {
  private client: MSGraphClient;
  private metadataExtensionName = 'mmsharepoint.onmicrosoft.MailStorage';
  private saveMetadata: boolean;

  constructor(saveMetadata: boolean) {
    this.saveMetadata = saveMetadata;
  }
  public init(graphFactory: MSGraphClientFactory): Promise<boolean> {
    return graphFactory
      .getClient()
      .then((client: MSGraphClient) => {
        this.client = client;
        return true;
      })
      .catch((error) => {
        return false;
      });
  }
  public getClient() {
    return this.client;
  }
  /**
   * This function retrieves all 1st-level folders from user's OneDrive
   */
  public getOneDriveFolder(): Promise<IFolder[]> {
    return this.client
      .api('me/drive/root/children')
      .version('v1.0')
      .filter('folder ne null')
      .select('id, name, parentReference, webUrl')
      .get()
      .then((response): any => {
        let folders: Array<IFolder> = new Array<IFolder>();
        response.value.forEach((item) => {
          folders.push({ id: item.id, name: item.name, driveID: item.parentReference.driveId, parentFolder: null, webUrl: item.webUrl });
        });
        return folders;
      });
  }
  public getGroupRootFolders(group: IFolder): Promise<IFolder[]> {
    return this.client
      .api(`drives/${group.driveID}/root/children`)
      .version('v1.0')
      .filter('folder ne null')
      .select('id, name, webUrl')
      .get()
      .then((response): any => {
        let folders: Array<IFolder> = new Array<IFolder>();
        response.value.forEach((item) => {
          folders.push({ id: item.id, name: item.name, driveID: group.driveID, parentFolder: group, webUrl: item.webUrl });
        });
        return folders;
      });
  }
  public getSubFolder(folder: any): Promise<any> {
    if (folder[0].id != folder[0].driveID) {
      return this.client
        .api(`drives/${folder[0].driveID}/items/${folder[0].id}/children`)
        .version('v1.0')
        .filter('folder ne null')
        .select('id,name,webUrl')
        .get()
        .then((response): any => {
          let folders: Array<any> = new Array<any>();
          response.value.forEach((item) => {
            folders.push({ id: item.id, name: item.name, driveID: folder[0].driveID, parentFolder: folder, webUrl: item.webUrl });
          });
          return folders;
        });
    }
    else {
      return this.client
        .api(`drives/${folder[0].driveID}/items/root/children`)
        .version('v1.0')
        .filter('folder ne null')
        .select('id,name,webUrl')
        .get()
        .then((response): any => {
          let folders: Array<any> = new Array<any>();
          response.value.forEach((item) => {
            folders.push({ id: item.id, name: item.name, driveID: folder[0].driveID, parentFolder: folder, webUrl: item.webUrl });
          });
          return folders;
        });
    }
  }
  /**Get site collection ID */
  public GetSiteID(siteName: string): Promise<any> {
    return this.client
      .api(`/sites/ccsdev01.sharepoint.com:/sites/${siteName}`)
      .version('v1.0')
      .select(['id'])
      .get().then(siteID => {
        console.log(siteID);
        return siteID;
      
      }, error => {
        return error;
      });
      
  }

  /**Get site collection ID */
  public savetoLog(siteID: string, mail: IMail, messages: any): Promise<any> {
    const metadataBody = {
      "fields": {
        "Title": mail.subject,
        "Description": messages,
        "LogType": "Mail"
      }
    };
    console.log("Inside log save", messages);
    return this.client
      .api(`https://graph.microsoft.com/v1.0/sites/${siteID}/lists/Log/items`)
      .post(JSON.stringify(metadataBody))
      .then((response) => {
        return 'Success';
      }).catch((error) => {
        return error.messages;
      });
  }

  /**
   * 
   * This function retrievs group id of site collection
   */
  public getGroupID(groupName: any): Promise<any> {
    return this.client
      .api('groups')
      .version('v1.0')
      .filter(`startswith(displayName, '${groupName}')`)
      .get()
      .then((response): any => {
        return response.value[0].id;
      });
  }
  /**
   * This function retrievs the user's membership groups from Graph
   */
  public getJoinedGroups(): Promise<IFolder[]> {
    return this.client
      .api('me/memberOf')
      .version('v1.0')
      .select('id, displayName, webUrl')
      .get()
      .then((response): any => {
        let folders: Array<IFolder> = new Array<IFolder>();
        response.value.forEach((item) => {
          // Show unified Groups but NO Teams
          if (item['@odata.type'] === '#microsoft.graph.group') {
            if (!item.resourceProvisioningOptions || item.resourceProvisioningOptions.indexOf('Team') === -1) {
              folders.push({ id: item.id, name: item.displayName, driveID: item.id, parentFolder: null, webUrl: item.webUrl });
            }
          }
        });
        return folders;
      });
  }

  /**
   * This function retrievs the user's membership groups from Graph
   */
  public getJoinedTeams(): Promise<IFolder[]> {
    return this.client
      .api('me/joinedTeams')
      .version('v1.0')
      .select('id, displayName, webUrl')
      .get()
      .then((response): any => {
        let folders: Array<IFolder> = new Array<IFolder>();
        response.value.forEach((item) => {
          folders.push({ id: item.id, name: item.displayName, driveID: item.id, parentFolder: null, webUrl: item.webUrl });
        });
        return folders;
      });
  }

  /**
   * This function retrieves all Drives for a given Group
   */
  public getGroupDrives(group: any): Promise<any> {
    return this.client
      .api(`groups/${group}/drives`)
      .version('v1.0')
      .select('id,name,webUrl')
      .get()
      .then((response): any => {
        let folders: Array<any> = new Array<any>();
        response.value.forEach((item) => {
          folders.push({ id: item.id, name: item.name, driveID: item.id, parentFolder: group, webUrl: item.webUrl });
        });
        return folders;
      });
  }

  public retrieveMimeMail = (driveID: string, folderID: string, mail: IMail, clientCallback: (msg: string) => void): Promise<string> => {
   
    return this.client
      .api(`me/messages/${mail.id}/$value`)
      .version('v1.0')
      .responseType('TEXT')
      .get()
      .then((response): any => {
        return response;
      });
  }
  //Retrive mail content
  public retrieveMailContent = (driveID: string, folderID: string, mail: IMail, clientCallback: (msg: string) => void): Promise<string> => {
    return this.client
      .api(`me/messages/${mail.id}`)
      .select('bodyPreview')
      .version('v1.0')
      .responseType('TEXT')
      .get()
      .then((response): any => {
        return JSON.parse(response);
      });
  }
  public saveNormalMail(driveID: string, folderID: string, mimeStream: string, fileName: string, clientCallback: (msg: string) => void): Promise<string> {
    const apiUrl = driveID !== folderID ? `drives/${driveID}/items/${folderID}:/${fileName}.eml:/content` : `drives/${driveID}/root:/${fileName}.eml:/content`;
    return this.client
      .api(apiUrl)
      .put(mimeStream)
      .then((response): any => {
        return response["webUrl"];
      })
      .catch((error) => {
        return error;
      });
  }

  public async saveBigMail(driveID: string, folderID: string, mimeStream: string, fileName: string, clientCallback: (msg: string) => void) {
    const sessionOptions = {
      "item": {
        "@microsoft.graph.conflictBehavior": "rename"
      }
    };
    const apiUrl = driveID !== folderID ? `drives/${driveID}/items/${folderID}:/${fileName}.eml:/createUploadSession` : `drives/${driveID}/root:/${fileName}.eml:/createUploadSession`;
    return this.client
      .api(apiUrl)
      .put(JSON.stringify(sessionOptions))
      .then(async (response): Promise<any> => {
        try {
          const resp = await this.uploadMailSlices(mimeStream, response.uploadUrl);
          return response["webUrl"];
        }
        catch (err) {
          return err;
        }
      });
  }

  private async uploadMailSlices(mimeStream: string, uploadUrl: string) {
    let minSize = 0;
    let maxSize = 5 * 327680; // 5*320kb slices --> MUST be a multiple of 320 KiB (327,680 bytes)
    while (mimeStream.length > minSize) {
      const fileSlice = mimeStream.slice(minSize, maxSize);
      const resp = await this.uploadMailSlice(uploadUrl, minSize, maxSize, mimeStream.length, fileSlice);
      minSize = maxSize;
      maxSize += 5 * 327680;
      if (maxSize > mimeStream.length) {
        maxSize = mimeStream.length;
      }
      if (resp.id !== undefined) {
        return resp;
      }
      else {

      }
    }
  }

  private async uploadMailSlice(uploadUrl: string, minSize: number, maxSize: number, totalSize: number, fileSlice: string) {
    const header = {
      "Content-Length": `${maxSize - minSize}`,
      "Content-Range": `bytes ${minSize}-${maxSize - 1}/${totalSize}`,
    };
    return await this.client
      .api(uploadUrl)
      .headers(header)
      .put(fileSlice);
  }

  public saveMailMetadata(mailId: string, displayName: string, url: string, savedDate: Date) {
    if (this.saveMetadata) {
      const apiUrl = `/me/messages/${mailId}/extensions`;
      const metadataBody = {
        "@odata.type": "microsoft.graph.openTypeExtension",
        "extensionName": this.metadataExtensionName,
        "saveDisplayName": displayName,
        "saveUrl": url,
        "savedDate": savedDate.toISOString()
      };
      this.client
        .api(apiUrl)
        .version('v1.0')
        .post(JSON.stringify(metadataBody))
        .then((response) => {
          console.log(response);
        });
    }
  }

  public retrieveMailMetadata(mailId: string): Promise<any> {
    const apiUrl = `/me/messages/${mailId}`;
    const expand = `Extensions($filter=id eq 'Microsoft.OutlookServices.OpenTypeExtension.${this.metadataExtensionName}')`;
    return this.client
      .api(apiUrl)
      .version('v1.0')
      .expand(expand)
      .select('id,subject,extensions')
      .get()
      .then((response) => {
        if (typeof response.extensions !== 'undefined' && response.extensions !== null) {
          const metadata: IMailMetadata = {
            extensionName: response.extensions[0].extensionName,
            saveDisplayName: response.extensions[0].saveDisplayName,
            saveUrl: response.extensions[0].saveUrl,
            savedDate: new Date(response.extensions[0].savedDate)
          };
          return metadata;
        }
        else {
          return null;
        }
      },
        (error) => {
          console.log(error);
          return null;
        });

  }

}
