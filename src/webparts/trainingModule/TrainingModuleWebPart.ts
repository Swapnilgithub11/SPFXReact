import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'TrainingModuleWebPartStrings';
import TrainingModule from './components/TrainingModule';
import { ITrainingModuleProps } from './components/ITrainingModuleProps';

export interface ITrainingModuleWebPartProps {
  description: string;
  //listname: string;
}

export default class TrainingModuleWebPart extends BaseClientSideWebPart<ITrainingModuleWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITrainingModuleProps > = React.createElement(
      TrainingModule,
      (() => {
       return {
          //spHttpClient: this.context.spHttpClient,
          description: this.properties.description,
          siteUrl: "https://jain1193.sharepoint.com",
          listname: "TrainingData",
          context: this.context,
          //FormDigestValue:'',
        };
      })()
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
                PropertyPaneTextField('listname', {
                  label: strings.ListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
