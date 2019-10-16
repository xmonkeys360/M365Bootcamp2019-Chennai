import { IGraphService } from './IGraphService';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClient } from '@microsoft/sp-http';

export class GraphService implements IGraphService {
    private context: WebPartContext;

    constructor(context: WebPartContext) {
        this.context = context;
    }

    public getGroupMembers(groupId: string): Promise<any[]> {
        console.log(groupId);
        if (groupId) {
            var result = new Promise<any[]>((resolve, reject) => {
                this.context.msGraphClientFactory
                    .getClient()
                    .then((graphClient: MSGraphClient): void => {
                        graphClient.api(`/groups/${groupId}/members`)
                            .get((error, data: any) => {
                                let members: any[] = [];
                                data.value.forEach((mem) => {
                                    members.push(mem);
                                });
                                console.log(members);
                                resolve(members);
                            });
                    });
            });
            return result;
        } else {
            console.log("error");
            return Promise.resolve();
        }
    }
}