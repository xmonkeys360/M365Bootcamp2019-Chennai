import * as React from 'react';
import styles from './SpoTeamsWebPart.module.scss';
import { ISpoTeamsWebPartProps } from './ISpoTeamsWebPartProps';
import { GraphService } from "../Services/GraphService";
import { SPService } from "../Services/SPService";
import TeamMembers from "./TeamMembers";
import {
  ListView,
  IViewField,
  SelectionMode,
} from "@pnp/spfx-controls-react/lib/ListView";
import { Stack } from 'office-ui-fabric-react/lib/Stack';

export interface ISpoTeamsWebPartState {
  members?: any[];
  selectedMember?: any;
  documents: any[];
  selectedDocument: any[];
}

export default class SpoTeamsWebPart extends React.Component<ISpoTeamsWebPartProps, ISpoTeamsWebPartState> {

  constructor(props: ISpoTeamsWebPartProps) {
    super(props);
    console.log(this.props);
    this.state = {
      members: [],
      selectedMember: null,
      documents: [],
      selectedDocument: []
    };
  }

  public componentDidMount() {
    console.log("component mount");
    let gs = new GraphService(this.props.spContext);
    let mems: any[] = [];
    gs.getGroupMembers(this.props.groupId).then(resp => {
      console.log(resp);
      mems = resp;
      this.setState({
        members: mems
      });
    });
  }

  private getSelectedMember(member: any) {
    console.log(member);
    this.setState({
      selectedMember: member
    });

    let sps = new SPService(this.props.spContext);
    let docs: any[] = [];
    sps.getMemberDocuments(member.mail, this.props.channelId).then(resp => {
      console.log(resp);
      this.setState({
        documents: resp
      });
    });
  }

  private _getSelection(items: any[]) {
    console.log("Selected items:", items);
    if (items.length > 0) {

      this.setState({
        selectedDocument: items
      });
    }
  }

  public render(): React.ReactElement<ISpoTeamsWebPartProps> {
    console.log(this.props);

    let title: string = "";
    let subTitle: string = "";
    let siteTabTitle: string = "";

    if (this.props.teamsContext) {
      title = "Welcome to teams!";
      subTitle = "Building custom enterprise tabs for your business.";
      siteTabTitle =
        "We are in the context of following Team: " + this.props.groupName;
    } else {
      title = "Welcome to SharePoint!";
      subTitle = "Customize SharePoint experiences using Web Parts.";
      siteTabTitle =
        "We are in the context of following site: " + this.props.spContext.pageContext.web.title;
    }

    const viewFields: IViewField[] = [
      {
        name: "Name"
      },
      {
        name: "Author"
      }
    ];

    return (
      <div className={styles.spoTeamsWebPart}>
        <div className="ms-Grid" dir="ltr">
          <div className={styles.toprow} >
            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg3"><span className={styles.rowtitle}>Group Members</span></div>
            <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6"><span className={styles.rowtitle}>Documents</span></div>
            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg3"></div>
          </div>
        </div>
        <div className="ms-Grid" dir="ltr">
          <div className={styles.fullrow} >
            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg3 rowsec">
              <div className={styles.rowsec}>
                <TeamMembers
                  members={this.state.members}
                  selectedMember={e => {
                    this.getSelectedMember(e);
                  }}
                />
              </div>
            </div>
            <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">
              <div className={styles.rowsec}>
                <Stack verticalFill={true}>
                  <ListView
                    items={this.state.documents}
                    viewFields={viewFields}
                    iconFieldName="ServerRelativeUrl"
                    compact={true}
                    selectionMode={SelectionMode.single}
                    showFilter={true}
                    selection={this._getSelection}
                    defaultFilter=""
                    filterPlaceHolder="Search..."
                  />
                </Stack>
              </div>
            </div>
            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg3">
              <div className={styles.container}>
                <div className={styles.row}>
                  <div className={styles.column}>
                    <span className={styles.title}>{title}</span>
                    <p className={styles.subTitle}>{subTitle}</p>
                    <p className={styles.description}>{siteTabTitle}</p>
                    <a href="https://aka.ms/spfx" className={styles.button}>
                      <span className={styles.label}>Learn more</span>
                    </a>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
