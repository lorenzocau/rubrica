import * as React from 'react';
import * as ReactDom from 'react-dom';

// per le chiamate a Microsoft Graph
import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'RubricaWebPartStrings';
import Rubrica from './components/Rubrica';
import { IRubricaProps } from './components/IRubricaProps';

export interface IRubricaWebPartProps {

}

export default class RubricaWebPart extends BaseClientSideWebPart<IRubricaWebPartProps> {



  public render(): void {
    const element: React.ReactElement<IRubricaProps> = React.createElement(
      Rubrica,
      // {}
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
