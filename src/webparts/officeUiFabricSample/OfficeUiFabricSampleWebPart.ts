import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'OfficeUiFabricSampleWebPartStrings';
import OfficeUiFabricSample from './components/OfficeUiFabricSample';
import { IOfficeUiFabricSampleProps } from './components/IOfficeUiFabricSampleProps';

export interface IOfficeUiFabricSampleWebPartProps {
  description: string;
}

export default class OfficeUiFabricSampleWebPart extends BaseClientSideWebPart<IOfficeUiFabricSampleWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IOfficeUiFabricSampleProps> = React.createElement(
      OfficeUiFabricSample,
      {
        description: this.properties.description,
        context:this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {

    return super.onInit();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
