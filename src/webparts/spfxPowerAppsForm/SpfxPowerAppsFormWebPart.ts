import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxPowerAppsFormWebPartStrings';
import SpfxPowerAppsForm from './components/SpfxPowerAppsForm';
import { ISpfxPowerAppsFormProps } from './components/ISpfxPowerAppsFormProps';

export interface ISpfxPowerAppsFormWebPartProps {
  description: string;
}

export default class SpfxPowerAppsFormWebPart extends BaseClientSideWebPart<ISpfxPowerAppsFormWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpfxPowerAppsFormProps > = React.createElement(
      SpfxPowerAppsForm,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
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
