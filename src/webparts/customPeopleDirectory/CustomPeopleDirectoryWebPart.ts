import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'CustomPeopleDirectoryWebPartStrings';
import CustomPeopleDirectory from './components/CustomPeopleDirectory';
import { ICustomPeopleDirectoryProps } from './components/ICustomPeopleDirectoryProps';
import { MSGraphClient } from "@microsoft/sp-http";


export interface ICustomPeopleDirectoryWebPartProps {
  description: string;
}

export default class CustomPeopleDirectoryWebPart extends BaseClientSideWebPart<ICustomPeopleDirectoryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICustomPeopleDirectoryProps > = React.createElement(
      CustomPeopleDirectory,
      {
        description: this.properties.description,
        graphClient : this.context.msGraphClientFactory
      }
    );

    ReactDom.render(element, this.domElement);
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
