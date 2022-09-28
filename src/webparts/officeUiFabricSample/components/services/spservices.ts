import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse , ISPHttpClientOptions} from '@microsoft/sp-http';
import { IDropdownOption } from "office-ui-fabric-react";
import { IUser } from "../interfaces/IUser";

export class SPServices {
    public GetAllLists(context:WebPartContext):Promise<IDropdownOption[]>
    {
        let restApiUri:string = context.pageContext.web.absoluteUrl + "/_api/web/lists?select=Title";
        var listTitles:IDropdownOption[]=[];
        return new Promise<IDropdownOption[]>(async(resolve,reject)=>{
            context.spHttpClient
            .get(restApiUri, SPHttpClient.configurations.v1)
            .then((response:SPHttpClientResponse)=>{
                response.json().then((results:any)=>{
                    results.value.map((result:any)=>{
                        listTitles.push({
                            key:result.Title,
                            text:result.Title
                        });
                    })
                });
                resolve(listTitles);
            }).catch((err:any)=>{
                reject("Error occurred: " + err);
                console.log(err);
            });
        })        
    }

    public GetListUsers(context:WebPartContext, listTitle:string):Promise<IDropdownOption[]>
    {
        let restApiUri:string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listTitle + "')/items";
        var userTitles:IDropdownOption[]=[];
        return new Promise<IDropdownOption[]>(async(resolve,reject)=>{
            context.spHttpClient
            .get(restApiUri, SPHttpClient.configurations.v1)
            .then((response:SPHttpClientResponse)=>{
                response.json()
                .then((results:any)=>{
                    var ids:number[] = [];
                    var uniqueIds:number[] = [];
                    var allIds:number[] = [];
                    for(var i in results.value)
                    {
                        allIds.push(results.value[i].UserId);
                        if(!ids.some(id => id === results.value[i].UserId))
                        {
                            ids.push(results.value[i].UserId);
                            uniqueIds.push(results.value[i].UserId);
                        }
                    }
                    for(var i in uniqueIds)
                    {
                        this.GetUserById(context, uniqueIds[i]).then((r:any)=>{
                            userTitles.push({
                                key:r.Email,
                                text:r.Title + " (" + r.Email + ")"
                            })
                        })
                    }      
                }).then((v)=>{
                    if(userTitles.some(v=>v.title === "adminjosh@4bbrwl.onmicrosoft.com"))
                    {
                        console.log("FOUND")
                    }
                });

                resolve(userTitles);
            }).catch((err:any)=>{
                reject("Error occurred: " + err);
                console.log(err);
            })
        })
    }

    public async GetListUsersAsync(context:WebPartContext, listTitle:string)
    {
        const listData = this.GetData(context, listTitle);
        listData.then((response:any)=>{
            let items:IDropdownOption[] = response.value;
            let dedup = items.filter((el, i, arr)=>{
                arr.indexOf(el) === i
            });
            return dedup;
        });
    }
    
    public GetUserById(context:WebPartContext, userId:number):Promise<IUser>
    {
        const restApiUri:string = context.pageContext.web.absoluteUrl + "/_api/web/getuserbyid(" + userId + ")";
        return new Promise<IUser>(async(resolve,reject)=>{
            context.spHttpClient
                .get(restApiUri, SPHttpClient.configurations.v1)
                .then((response:SPHttpClientResponse)=>{
                    response.json().then((results:any)=>{
                        resolve(results); 
                    });
                }).catch((err:any)=>{
                    reject("Error occured: " + err);
                    console.log(err);
                })
        })
    }

    public async GetUserByIdAsync(context:WebPartContext, userId:number)
    {
        const restApiUri:string = context.pageContext.web.absoluteUrl + "/_api/web/getuserbyid(" + userId + ")";
        const response = await context.spHttpClient.get(restApiUri, SPHttpClient.configurations.v1);
        return await response.json();   
    }

    public GetUserByEmail(context:WebPartContext, userEmail:string):Promise<IUser>
    {
        const restApiUri:string = context.pageContext.web.absoluteUrl + "/_api/web/SiteUsers?$filter=Email eq '" + userEmail + "'";
        return new Promise<IUser>(async(resolve,reject)=>{
            context.spHttpClient
                .get(restApiUri, SPHttpClient.configurations.v1)
                .then((response:SPHttpClientResponse)=>{
                    response.json().then((results:any)=>{
                        console.log(results);
                        resolve(results); 
                    });
                }).catch((err:any)=>{
                    reject("Error occured: " + err);
                    console.log(err);
                })
        })
    }

    public async GetData(context:WebPartContext, listTitle:string):Promise<SPHttpClientResponse>
    {
        const restApiUri:string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listTitle + "')/items";
        const response = await context.spHttpClient.get(restApiUri, SPHttpClient.configurations.v1);
        const result = response.json();
        return await result;
    }
}