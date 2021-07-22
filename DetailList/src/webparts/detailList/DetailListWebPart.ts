import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DetailListWebPartStrings';
import DetailList from './components/DetailList';
import { IDetailListProps } from './components/IDetailListProps';
import { sp } from '@pnp/sp';

export interface IDetailListWebPartProps {
  description: string;
  siteurl:string;
  listName:string;
}

export default class DetailListWebPart extends BaseClientSideWebPart<IDetailListWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IDetailListProps> = React.createElement(
      DetailList,
      {
        description: this.properties.description,
        siteurl: this.properties.siteurl,
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
                PropertyPaneTextField('siteurl', {
                  label: "Siteurl"
                }),
                PropertyPaneTextField('listName', {
                  label: "List Name"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
