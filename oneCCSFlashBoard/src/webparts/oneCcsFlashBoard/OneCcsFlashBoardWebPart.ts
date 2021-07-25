import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneTextField, IPropertyPaneCheckboxProps
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'OneCcsFlashBoardWebPartStrings';
import OneCcsFlashBoard from './components/OneCcsFlashBoard';
import { IOneCcsFlashBoardProps } from './components/IOneCcsFlashBoardProps';
import $ from 'jquery';
import { sp } from "@pnp/sp/presets/all";
export interface IOneCcsFlashBoardWebPartProps {
  description: string;
}

export default class OneCcsFlashBoardWebPart extends BaseClientSideWebPart<IOneCcsFlashBoardProps> {
  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      sp.setup({
        spfxContext: this.context,
      });
    });
  }
  public render(): void {
    try {

      $(".ControlZone").parent().parent().css("max-width", "100%");
    }
    catch (err) {
      console.log("Couldnot update the max-width of the page");
    }
    const element: React.ReactElement<IOneCcsFlashBoardProps> = React.createElement(
      OneCcsFlashBoard,
      {
        description: this.properties.description,
        listName: this.properties.listName,
        seperator: this.properties.seperator,
        speed: this.properties.speed,
        siteUrl: this.context.pageContext.web.serverRelativeUrl,
        horizontal: this.properties.horizontal,
        vertical: this.properties.vertical,

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
                PropertyPaneTextField('listName', {
                  label: "ListName"
                }),
                PropertyPaneTextField('seperator', {
                  label: "Seperator"
                }),
                PropertyPaneTextField('speed', {
                  label: "Speed"
                }),
                PropertyPaneCheckbox('horizontal', {
                  text: "Horizontal"
                }),
                PropertyPaneCheckbox('vertical', {
                  text: "Vertical"
                }),
                PropertyPaneCheckbox('slider', {
                  text: "Carosal"
                }),

              ]
            }
          ]
        }
      ]
    };
  }
}
