import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'RevisionHistoryWebPartStrings';
import RevisionHistory from './components/RevisionHistory';
import { IRevisionHistoryProps } from './components/IRevisionHistoryProps';
import { sp } from "@pnp/sp/presets/all";

export interface IRevisionHistoryWebPartProps {
  description: string;

}

export default class RevisionHistoryWebPart extends BaseClientSideWebPart<IRevisionHistoryProps> {
  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      // other init code may be present
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IRevisionHistoryProps> = React.createElement(
      RevisionHistory,
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
                  label: strings.SiteUrlFieldLabel
                }),
                PropertyPaneTextField('listName', {
                  label: strings.listName
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
