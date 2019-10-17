import * as React from 'react';
import styles from './Graphpersona.module.scss';
import { IGraphpersonaProps } from './IGraphpersonaProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { IGraphpersonaState } from './IGraphpersonaState';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

//import {Client} from '@microsoft/microsoft-graph-client';

// import {
//   Persona,
//   PersonaSize
// } from 'office-ui-fabric-react/lib/components/Persona';
// import { Link } from 'office-ui-fabric-react/lib/components/Link';

export default class Graphpersona extends React.Component<IGraphpersonaProps, IGraphpersonaState> {

  constructor(props: IGraphpersonaProps) {
    super(props);
  
    this.state = {
      name: '',
      email: '',
      phone: '',
      image: null
       
    };
  }

  public componentDidMount(): void {
    this.props.graphClient
      .api('me')
      .get((error: any, user: MicrosoftGraph.User, rawResponse?: any) => {
        this.setState({
          name: user.displayName,
          email: user.mail,
          phone: user.businessPhones[0]
           
        });
      });
  
    this.props.graphClient
      .api('/me/photo/$value')
      .responseType('blob')
      .get((err: any, photoResponse: any, rawResponse: any) => {
        const blobUrl = window.URL.createObjectURL(rawResponse.xhr.response);
        this.setState({ image: blobUrl });
      });
  }

  private _renderMail = () => {
    if (this.state.email) {
      //  return <Link href={`mailto:${this.state.email}`}>{this.state.email}</Link>;
    } else {
      return <div />;
    }
  }
  
  private _renderPhone = () => {
     if (this.state.phone) {
      //  return <Link href={`tel:${this.state.phone}`}>{this.state.phone}</Link>;
    } else {
      return <div />;
    }
  }


  public render(): React.ReactElement<IGraphpersonaProps> {
    return (
      // <Persona primaryText={this.state.name}
      //         secondaryText={this.state.email}
      //         onRenderSecondaryText={this._renderMail}
      //         tertiaryText={this.state.phone}
      //         onRenderTertiaryText={this._renderPhone}
      //         imageUrl={this.state.image}
      //         size={PersonaSize.size100} />
      <div>this.state.name</div>
    );
  }
}
