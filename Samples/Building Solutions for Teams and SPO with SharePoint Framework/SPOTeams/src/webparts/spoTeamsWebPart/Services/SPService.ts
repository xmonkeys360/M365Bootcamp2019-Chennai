import {ISPService} from './ISPService';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

export class SPService implements ISPService{
    private context: WebPartContext;

    constructor(context:WebPartContext){
        this.context = context;
    }

    public getMemberDocuments(emailid:string, libTitle:string):Promise<any[]>{
        if(emailid){
            var result = new Promise<any[]>((resolve,reject)=>{

                const siteUrl  = this.context.pageContext.web.absoluteUrl;
                this.context.spHttpClient
                .fetch(`${siteUrl}/_api/lists/GetByTitle('Documents')/items?$filter=Author/EMail eq '${emailid}'&$select=Title,FileLeafRef,FileRef,UniqueId,Modified,Author/Name,Author/Title,Author/EMail&$expand=Author/Id&$orderby=Title`,
                SPHttpClient.configurations.v1,                
                    {
                        method: 'GET',
                        headers: { "accept": "application/json" },
                        mode: 'cors',
                        cache: 'default'
                    }                
                )
                 .then((response) => {
                    if (response.ok) {
                        return response.json();
                    } else {
                        throw (`Error ${response.status}: ${response.statusText}`);
                    }
                })
                .then((o: any) => {
                    console.log(o);
                    let docs: any[] = [];
                    o.value.forEach((doc) => {

                        let ext = doc.FileLeafRef.split('.');

                        docs.push({
                            Name: doc.FileLeafRef,
                            Path: doc.FileRef,
                            Author: doc.Author.Title,
                            Modified: new Date(doc.Modified)                            
                        });
                    });
                    resolve(docs);
                });
            });
            return result;
        }else{
            return Promise.resolve();
        }
    }
}