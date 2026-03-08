import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import Rubrica from './components/Rubrica';
import { Sharepoint } from '@microsoft/microsoft-graph-types';

export default class RubricaWebPart extends BaseClientSideWebPart<{}> {

  public render(): void {
    const element: React.ReactElement<IRubricaProps> = React.createElement(
      Rubrica,
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [] };
  }
}