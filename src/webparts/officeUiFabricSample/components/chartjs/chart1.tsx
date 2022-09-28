import * as React from 'react';
import { Chart, registerables } from 'chart.js';
import { Pie } from 'react-chartjs-2';
import { IOfficeUiFabricSampleState } from '../IOfficeUiFabricSampleState';
Chart.register(...registerables);

export function Chart1(state:IOfficeUiFabricSampleState) {
    return <Pie 
        data={{
        labels: state.chartData.labels,
        datasets: [
            {
                label: state.chartData.label,
                data: state.chartData.data,
                backgroundColor: state.chartData.backgroundColor,
                borderColor: state.chartData.borderColor,
                borderWidth: state.chartData.borderWidth,
            },
        ],
        }} 
    />;
}
