import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IDocumentReviewProps {
  description: string;
  context: WebPartContext;
  project:string;
}
