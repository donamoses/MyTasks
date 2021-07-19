import GraphController from '../../../controller/GraphController';
import { IMailMetadata } from '../../../model/IMailMetadata';
import { IFolder } from '../../../model/IFolder';
export interface IOutlook2SharePointState {
  graphController: GraphController;
  mailMetadata: IMailMetadata;
  showSuccess: boolean;
  showError: boolean;
  showOneDrive: boolean;
  showTeams: boolean;
  showGroups: boolean;
  successMessage: string;
  errorMessage: string;
  selectedHubOptions: string;
  hubsiteSelected: any;
  listSelected: any;
  siteSelected: any;
  folderStructureItems: string[];
  currentRootFolder: any;
  key: number;
  initialFocusedIndex?: number;
  folders: IFolder[];
  grandParentFolder: IFolder;
  parentFolder: any;
  selectedGroupName: string;
  showSpinner: boolean;
  selectedFolder: any;
  hideLoader: string;
}
