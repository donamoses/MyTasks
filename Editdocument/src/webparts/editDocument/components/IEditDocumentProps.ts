import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IEditDocumentProps {
  context: WebPartContext;
  description: string;
  createdocument:string;
  project:string;
}
