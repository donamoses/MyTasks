import { MSGraphClientFactory } from '@microsoft/sp-http';
import { IMail } from '../../../model/IMail';
import GraphController from '../../../controller/GraphController';
export interface IOutlook2SharePointProps {
  mail: IMail;
  msGraphClientFactory: MSGraphClientFactory;
  webServerRelativeUrl: string;
  graphController: GraphController;
  successCallback: (msg: string) => void;
  errorCallback: (msg: string) => void;
}
