import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { sp } from '@pnp/sp';
import * as strings from 'DocumentReviewWebPartStrings';
import DocumentReview from './components/DocumentReview';
import { IDocumentReviewProps } from './components/IDocumentReviewProps';

export interface IDocumentReviewWebPartProps {
  description: string;
  project:string;
}

export default class DocumentReviewWebPart extends BaseClientSideWebPart<IDocumentReviewWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IDocumentReviewProps> = React.createElement(
      DocumentReview,
      {
        context: this.context,
        description: this.properties.description,
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
