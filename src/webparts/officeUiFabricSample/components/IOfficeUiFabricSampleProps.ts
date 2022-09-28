import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IOfficeUiFabricSampleState } from "./IOfficeUiFabricSampleState";

export interface IOfficeUiFabricSampleProps {
  description: string;
  context:WebPartContext;
}
