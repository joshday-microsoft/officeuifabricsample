import * as React from 'react';
import { IOfficeUiFabricSampleProps } from './IOfficeUiFabricSampleProps';
import { IOfficeUiFabricSampleState } from './IOfficeUiFabricSampleState';
import { Main } from './fluentui/main';
import { SPServices } from './services/spservices';
import { IDropdownOption } from 'office-ui-fabric-react';
import HelperService from './services/helperservice';
import ChartService from './services/chartservice';
import { initializeIcons } from '@fluentui/react/lib/Icons';

initializeIcons(/* optional base url */);
export default class OfficeUiFabricSample extends React.Component<IOfficeUiFabricSampleProps, IOfficeUiFabricSampleState, {}> {
  public _spServices: SPServices;
  public _helperService: HelperService;
  public _chartService: ChartService;

  constructor(props:IOfficeUiFabricSampleProps){
    super(props);
    this._spServices = new SPServices();
    this._helperService = new HelperService(props);
    this._chartService = new ChartService(props);
    
    this.state={
      listTitles: [], 
      chartData:null,
      listItems:[],
      listUsers:[],
      selectedUser: {
        key: "",
        text: ""
      },
      handleUserChange: ()=>{null}, 
      handleClick: ()=>{null}
    };
  }

  public componentDidMount(): void {
    this._spServices.GetAllLists(this.props.context)
        .then((result:IDropdownOption[])=>{
        this.setState({listTitles:result});
      });

    this._spServices.GetListUsers(this.props.context, "File Audit Log")
      .then((result:IDropdownOption[])=>{
      this.setState({listUsers:result});
    });
    
    this.setState({handleUserChange: (element:any, func:any) => this.handleUserChange(element, func)})
    this.setState({handleClick: ()=> this.handleClick()});
  }

  public handleUserChange(element:any, func:any)
  {
    this.setState({selectedUser: {
      key:func.key,
      text:func.text
    }});

    this._chartService.getChartData(func.key).then((result:any)=>{
      this.setState({chartData:result});
    });
  }

  public handleClick()
  {
    this.refreshData();
  }

  public refreshData()
  {
    this._spServices.GetAllLists(this.props.context)
    .then((result:IDropdownOption[])=>{
      this.setState({listTitles:result});
    });

  this._spServices.GetListUsers(this.props.context, "File Audit Log")
    .then((result:IDropdownOption[])=>{
      this.setState({listUsers:result});
    });

  this._chartService.getChartData(this.state.selectedUser.key.toString()).then((result:any)=>{
    this.setState({chartData:result});
  });
  }

  public render(): React.ReactElement<IOfficeUiFabricSampleProps> {
    
    return (
        <Main
          selectedUser={this.state.selectedUser}
          handleUserChange={this.state.handleUserChange}
          listUsers={this.state.listUsers} 
          listTitles={this.state.listTitles} 
          listItems={this.state.listItems} 
          chartData={this.state.chartData}
          handleClick={this.state.handleClick}
        />
    );
  }
}
