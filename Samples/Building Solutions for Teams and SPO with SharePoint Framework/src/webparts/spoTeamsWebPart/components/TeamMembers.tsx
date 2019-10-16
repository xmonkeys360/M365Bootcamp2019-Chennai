import * as React from 'react';
import styles from './SpoTeamsWebPart.module.scss';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';

export interface ITeamMembersProps {
    members: any[];
    selectedMember: any;
}
export interface ITeamMembersState {
    users: any[];
}
export default class TeamMembers extends React.Component<ITeamMembersProps, ITeamMembersState> {

    constructor(props: ITeamMembersProps) {
        super(props);
        console.log(this.props);
        this.handleMoreInfo = this.handleMoreInfo.bind(this);
        this.state = {
            users: this.props.members.length > 0 ? this.props.members : []
        };
    }

    private handleMoreInfo(data: any, usersList: any) {
        console.log(data);
        console.log(usersList);
        this.props.selectedMember(data);
    }



    public render(): React.ReactElement<ITeamMembersProps> {

        let persComp = this.props.members.length > 0 ? this.props.members.map((x, i) => {
            return <Persona
                className={styles.stackPersona}

                size={PersonaSize.size32}
                imageUrl={`/_layouts/15/userphoto.aspx?UA=0&size=HR64x64&accountname=` + x.mail}
                text={x.displayName}
                secondaryText={x.mail}
                onClick={this.handleMoreInfo.bind(this, x)}

            />

        }) : '';
        return (

            <Stack className={styles.stackMembers} verticalFill={true} tokens={{ childrenGap: 15 }}>
                {persComp}
            </Stack>
        );
    }
}