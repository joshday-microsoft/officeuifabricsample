import * as React from 'react';
import { IOfficeUiFabricSampleProps } from '../IOfficeUiFabricSampleProps';
import { IOfficeUiFabricSampleState } from '../IOfficeUiFabricSampleState';
import ChartService from './chartservice';
import { SPServices } from './spservices';

export default class HelperService extends React.Component<IOfficeUiFabricSampleProps, IOfficeUiFabricSampleState, {}> {
    public _spServices: SPServices;
    public _chartService: ChartService;

    constructor(props:IOfficeUiFabricSampleProps){
        super(props);
        this._spServices = new SPServices();
        this._chartService = new ChartService(props);
    }

    componentDidMount(): void {
        
    }
    
    public GenerateRandomNumber():number
    {
        var min:number = Math.ceil(0);
        var max:number = Math.floor(255);
        return Math.floor(Math.random() * (max - min + 1)) + min;
    }
}