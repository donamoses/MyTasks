import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { sp } from "@pnp/sp/presets/all";
import * as strings from 'EditDocumentWebPartStrings';
import EditDocument from './components/EditDocument';
import { IEditDocumentProps } from './components/IEditDocumentProps';

export interface IEditDocumentWebPartProps {
  description: string;
  createdocument:string;
  project:string;
}

export default class EditDocumentWebPart extends BaseClientSideWebPart<IEditDocumentWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IEditDocumentProps> = React.createElement(
      EditDocument,
      {
        context: this.context,
        description: this.properties.description,
        createdocument:this.properties.createdocument,
        project:this.properties.project
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneToggle('createdocument',{
                  label:'CreateDocument',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneToggle('project',{
                  label:'Project',
                  onText: 'On',
                  offText: 'Off'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
