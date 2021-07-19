import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface ICreateDocumentProps {
  description: string;
  EmployeelistName: string;
  siteUrl: string;
  context: WebPartContext;
  RedirectUrl:string;
  EmployeeUrl:string;
  DepartmentlistName: string;
  DocumentlistName: string;
  TemplatelistName:string;
  TemplateCategory:string;
  ListName:string;
}
