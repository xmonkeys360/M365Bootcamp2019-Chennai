export interface IGraphService{
    getGroupMembers(groupId:string):Promise<any[]>;    
}