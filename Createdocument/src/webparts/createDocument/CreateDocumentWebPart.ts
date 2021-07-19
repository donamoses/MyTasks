import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CreateDocumentWebPartStrings';
import CreateDocument from './components/CreateDocument';
import { ICreateDocumentProps } from './components/ICreateDocumentProps';
import { sp } from "@pnp/sp/presets/all";
export interface ICreateDocumentWebPartProps {
  description: string;
  siteUrl: string;
  RedirectUrl:string;
  EmployeeUrl:string;
  EmployeelistName: string;
  DepartmentlistName: string;
  DocumentlistName: string;
  TemplatelistName:string;
  TemplateCategory:string;
  ListName:string;
}

export default class CreateDocumentWebPart extends BaseClientSideWebPart<ICreateDocumentWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<ICreateDocumentProps> = React.createElement(
      CreateDocument,
      {
        context: this.context,
        description: this.properties.description,
        EmployeelistName: this.properties.EmployeelistName,
        siteUrl: this.properties.siteUrl,
        RedirectUrl:this.properties.RedirectUrl,
        EmployeeUrl:this.properties.EmployeeUrl,
        DepartmentlistName:this.properties.DepartmentlistName,
        DocumentlistName:this.properties.DocumentlistName,
        TemplatelistName:this.properties.TemplatelistName,
        TemplateCategory:this.properties.TemplateCategory,
        ListName:this.properties.ListName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                 PropertyPaneTextField('siteUrl', {
                  label: "SiteUrl"
                }),
                  PropertyPaneTextField('RedirectUrl', {
                  label: "Redirect URL"
                }),
                PropertyPaneTextField('EmployeeUrl', {
                  label: "Employee URL"
                }),
                PropertyPaneTextField('EmployeelistName', {
                  label: "Employee ListName"
                }),
                PropertyPaneTextField('DepartmentlistName', {
                  label: "Department ListName"
                }),
                PropertyPaneTextField('DocumentlistName', {
                  label: "Document ListName"
                }),
                PropertyPaneTextField('TemplatelistName', {
                  label: "Template ListName"
                }),
                PropertyPaneTextField('TemplateCategory', {
                  label:"Template Category"
                }),
                PropertyPaneTextField('ListName', {
                  label:"ListName"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
