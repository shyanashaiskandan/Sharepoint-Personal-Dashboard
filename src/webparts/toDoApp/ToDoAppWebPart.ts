import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import ToDoApp from './components/ToDoApp';
import { ToDoAppProps } from './components/ToDoAppProps';

export interface IToDoAppWebPartProps {
  description: string;
}

export default class ToDoAppWebPart extends BaseClientSideWebPart<IToDoAppWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ToDoAppProps> = React.createElement(ToDoApp, {
      description: this.properties.description,
      context: this.context
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

}