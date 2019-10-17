import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';


import * as strings from 'GraphpersonaWebPartStrings';
import Graphpersona from './components/Graphpersona';
import { IGraphpersonaProps } from './components/IGraphpersonaProps';

import { MSGraphClient } from '@microsoft/sp-http';
 

export interface IGraphpersonaWebPartProps {
  description: string;
}

export default class GraphpersonaWebPart extends BaseClientSideWebPart<IGraphpersonaWebPartProps> {

  public render(): void {
    this.context.msGraphClientFactory.getClient()
    .then((client: MSGraphClient): void => {
      const element: React.ReactElement<IGraphpersonaProps> = React.createElement(
        Graphpersona,
        {
          graphClient: client
        }
      );
      
      ReactDom.render(element, this.domElement);
    });    
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
