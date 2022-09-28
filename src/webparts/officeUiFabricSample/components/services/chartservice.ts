import { SPServices } from "./spservices";
import { IListItem } from '../interfaces/IListItem';
import { IOfficeUiFabricSampleProps } from "../IOfficeUiFabricSampleProps";
import * as React from "react";
import { IChartData } from "../interfaces/IChartData";
import { IUser } from "../interfaces/IUser";
import { IChartProps } from "../interfaces/IChartProps";
import { IItemCount } from "../interfaces/IItemCount";

export default class ChartService extends React.Component<IOfficeUiFabricSampleProps> {
    public _spServices: SPServices;

    constructor(props:IOfficeUiFabricSampleProps)
    {
        super(props);
        this._spServices = new SPServices();   
    }

    public async getChartProps(userEmail:string):Promise<IChartProps>
    {
        let chartProps:IChartProps = {
            uniqueFilteredLabels: [],
            allFilteredLabels: [],
            data:[],
            borderColorArr:[],
            backgroundColorArr:[]
        }
        var labelsArr:Array<string>=[];
        var uniqueFilteredLabelArr:Array<string>=[];
        var allFilteredLabels:Array<string>=[];
        var borderColorArr:Array<string>=[];
        var backgroundColorArr:Array<string>=[];
        var itemCountArr:IItemCount[]=[];
        var dataArr:number[]=[];

        await this._spServices.GetData(this.props.context, "File Audit Log").then(async (response:any)=>{
            var items:IListItem[] = response.value;
            for(var i in items)
            {
                var email = await this._spServices.GetUserById(this.props.context, items[i].UserId)
                    .then((user:IUser)=>{
                        return user.Email;
                    });
                if(email.toLowerCase() == userEmail.toLowerCase())
                {
                    allFilteredLabels.push(items[i].FileAccessed);
                    if(!labelsArr.some(label => label.toLowerCase() === items[i].FileAccessed.toLowerCase()))
                    {   
                        labelsArr.push(items[i].FileAccessed);
                        uniqueFilteredLabelArr.push(items[i].FileAccessed);
                        var num1 = this.GenerateRandomNumber();
                        var num2 = this.GenerateRandomNumber();
                        var num3 = this.GenerateRandomNumber();

                        itemCountArr.push({
                            File: items[i].FileAccessed,
                            Count: 0
                        });
                        backgroundColorArr.push('rgba('+num1+', '+num2+', '+num3+', 0.2)');
                        borderColorArr.push('rgba('+num1+', '+num2+', '+num3+', 1)')
                    }
                }
            }            

            //START: Calculate Data
            for(var i in allFilteredLabels)
            {   
                for(var j in itemCountArr)
                {
                    if(allFilteredLabels[i]===itemCountArr[j].File)
                    {
                        itemCountArr[j].Count = itemCountArr[j].Count + 1
                    }
                }
            }
            
            
            for(var i in itemCountArr)
            {
                dataArr.push(itemCountArr[i].Count)
            }
            //END: Calculate Data

            chartProps.data = dataArr;
            chartProps.uniqueFilteredLabels = uniqueFilteredLabelArr;
            chartProps.allFilteredLabels = allFilteredLabels;
            chartProps.backgroundColorArr = backgroundColorArr;
            chartProps.borderColorArr = borderColorArr;
        });
        
        return chartProps;
    }

    public async getChartData(user:string):Promise<IChartData>
    {
        let l:IChartProps = await this.getChartProps(user);
        let chartData:IChartData = {
            data: l.data,
            label: "Files Accessed by " + user,
            labels: l.uniqueFilteredLabels,
            backgroundColor: l.backgroundColorArr,
            borderColor: l.borderColorArr,
            borderWidth: 1
        }
        return chartData;
    }

    public GenerateRandomNumber():number
    {
        var min:number = Math.ceil(0);
        var max:number = Math.floor(255);
        return Math.floor(Math.random() * (max - min + 1)) + min;
    }
}