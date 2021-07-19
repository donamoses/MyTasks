import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface ISendRequestProps {
  description: string;
  context: WebPartContext;
  project:string;
}
