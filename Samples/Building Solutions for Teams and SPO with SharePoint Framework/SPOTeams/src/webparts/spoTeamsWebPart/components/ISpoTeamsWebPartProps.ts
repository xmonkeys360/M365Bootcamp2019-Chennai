import * as microsoftTeams from '@microsoft/teams-js';
import {  WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISpoTeamsWebPartProps {
  spContext?:WebPartContext;
  teamsContext?:microsoftTeams.Context;
  groupName?:string;
  groupId?:string;
  channelId?:string;
}
