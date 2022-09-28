import * as React from 'react';
import { Chart1 } from '../chartjs/chart1';
import { IOfficeUiFabricSampleState } from '../IOfficeUiFabricSampleState';
import { Dropdown, ActionButton, IIconProps } from 'office-ui-fabric-react';

export function Main(state:IOfficeUiFabricSampleState) {
    const refreshIcon: IIconProps = {iconName: 'Refresh'}
    return (
        <div>
            <ActionButton iconProps={refreshIcon} onClick={()=>state.handleClick()}>Refresh</ActionButton>
            <Dropdown label="Choose a User" options={state.listUsers} onChange={state.handleUserChange}/>
            <br/>
            {
                state.chartData ? 
                    <>
                        <Chart1 handleClick={state.handleClick} handleUserChange={state.handleUserChange} selectedUser={state.selectedUser} listUsers={state.listUsers} listTitles={state.listTitles} listItems={state.listItems} chartData={state.chartData} /> 
                    </>
                    : 
                    null
            }

        </div>
    );
}