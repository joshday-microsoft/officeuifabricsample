import { IDropdownOption } from "office-ui-fabric-react";
import { IChartData } from "./interfaces/IChartData";
import { IListItem } from "./interfaces/IListItem";

export interface IOfficeUiFabricSampleState {
  listTitles: IDropdownOption[];
  listUsers: IDropdownOption[];
  listItems: IListItem[];
  chartData: IChartData;
  selectedUser: IDropdownOption;
  handleUserChange: (element:any, func:any) => void;
  handleClick: () => void;
}
