import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'OneCcsApplicationDashboardWebPartStrings';
import OneCcsApplicationDashboard from './components/OneCcsApplicationDashboard';
import { IOneCcsApplicationDashboardProps } from './components/IOneCcsApplicationDashboardProps';
import { sp } from "@pnp/sp/presets/all";

export interface IOneCcsApplicationDashboardWebPartProps {
  description: string;
}

export default class OneCcsApplicationDashboardWebPart extends BaseClientSideWebPart<IOneCcsApplicationDashboardProps> {
  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      // other init code may be present
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IOneCcsApplicationDashboardProps> = React.createElement(
      OneCcsApplicationDashboard,
      {
        description: this.properties.description,
        siteUrl: this.properties.siteUrl,
        listName: this.properties.listName,
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
                PropertyPaneTextField('siteUrl', {
                  label: "siteUrl"
                }),
                PropertyPaneTextField('listName', {
                  label: "listName"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
