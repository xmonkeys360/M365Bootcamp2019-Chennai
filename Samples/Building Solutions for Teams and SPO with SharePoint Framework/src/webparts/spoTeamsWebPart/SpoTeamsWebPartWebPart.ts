import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'SpoTeamsWebPartWebPartStrings';
import SpoTeamsWebPart from './components/SpoTeamsWebPart';
import { ISpoTeamsWebPartProps } from './components/ISpoTeamsWebPartProps';

import * as microsoftTeams from '@microsoft/teams-js';



export interface ISpoTeamsWebPartWebPartProps {
  groupName: string;
  groupId: string;
  channelId: string;
}

export default class SpoTeamsWebPartWebPart extends BaseClientSideWebPart<ISpoTeamsWebPartWebPartProps> {

  private _spContext: WebPartContext;
  private _teamsContext: microsoftTeams.Context;
  private groupName?: string;
  private groupId?: string;
  private channelId?: string;

  protected onInit(): Promise<any> {

    let p: Promise<any> = Promise.resolve();

    if (this.context.microsoftTeams &&
      this.context.microsoftTeams.getContext) {

      // Get configuration from the Teams Client SDK
      p = new Promise((resolve, reject) => {
        if (this.context.microsoftTeams &&
          this.context.microsoftTeams.getContext) {
          this.context.microsoftTeams.getContext(context => {

            this._teamsContext = context;
            this.groupName = context.teamName;
            this.groupId = context.groupId;
            this.channelId = context.channelId;
            resolve();
          });
        }
      });

    } else {

      // Get configuration from web part settings from SharePoint
      this.groupName = this.properties.groupName;
      this.groupId = this.properties.groupId;
      this.channelId = this.properties.channelId;
    }
    return p;
  }

  public async render(): Promise<void> {
    this._spContext = this.context;
    await import('./styles/maxWidth.module.scss');
    const element: React.ReactElement<ISpoTeamsWebPartProps> = React.createElement(
      SpoTeamsWebPart,
      {
        spContext: this._spContext,
        teamsContext: this._teamsContext,
        groupName: this.groupName,
        groupId:this.groupId,
        channelId:this.channelId
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
                PropertyPaneTextField('groupId', {
                  label: strings.GidFieldLabel
                }),
                PropertyPaneTextField('groupName', {
                  label: strings.GNameFieldLabel
                }),
                PropertyPaneTextField('channelId', {
                  label: strings.CidFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
