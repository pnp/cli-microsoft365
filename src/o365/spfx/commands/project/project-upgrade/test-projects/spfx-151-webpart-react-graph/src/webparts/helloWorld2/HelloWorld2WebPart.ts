import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'HelloWorld2WebPartStrings';
import HelloWorld2 from './components/HelloWorld2';
import { IHelloWorld2Props } from './components/IHelloWorld2Props';
import { MSGraphClient } from '@microsoft/sp-client-preview';

export interface IHelloWorld2WebPartProps {
  description: string;
}

export default class HelloWorld2WebPart extends BaseClientSideWebPart<IHelloWorld2WebPartProps> {

  public render(): void {
    const graphClient: MSGraphClient = this.context.serviceScope.consume(MSGraphClient.serviceKey);
    const element: React.ReactElement<IHelloWorld2Props > = React.createElement(
      HelloWorld2,
      {
        description: this.properties.description,
        graphClient: graphClient
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
