import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IDocumentApprovalProps {
  description: string;
  project: string;
  context: WebPartContext;
}
