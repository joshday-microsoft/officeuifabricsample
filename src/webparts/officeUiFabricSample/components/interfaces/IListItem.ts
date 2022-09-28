import { IUser } from "./IUser"

export interface IListItem {
    Id:number,
    Title:string,
    User: IUser,
    UserId: number,
    FileAccessed:string,
    DateAccessed:Date
}