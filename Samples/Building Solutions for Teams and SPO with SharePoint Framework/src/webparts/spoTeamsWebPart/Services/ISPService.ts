export interface ISPService{
    getMemberDocuments(emailid:string,libTitle:string):Promise<any[]>;
}