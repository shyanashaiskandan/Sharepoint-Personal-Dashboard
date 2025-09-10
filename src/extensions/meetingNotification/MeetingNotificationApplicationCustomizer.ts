import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import Banner from './components/Banner';

export default class MeetingNotificationApplicationCustomizer
  extends BaseApplicationCustomizer<{}> {

  private _topPlaceholder?: PlaceholderContent;

  public async onInit(): Promise<void> {
    this._renderTop();
    return;
  }

  private _renderTop(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        {
          onDispose: () => {
            if (this._topPlaceholder?.domElement) {
              ReactDOM.unmountComponentAtNode(this._topPlaceholder.domElement);
            }
          }
        }
      );
    }
    if (!this._topPlaceholder?.domElement) return;

    const onDismiss = () => {
      if (this._topPlaceholder?.domElement) {
        ReactDOM.unmountComponentAtNode(this._topPlaceholder.domElement);
      }
    };

    ReactDOM.render(
      React.createElement(Banner, {
        message: "Your meeting starts in 5 minutes!",
        onDismiss
      }),
      this._topPlaceholder.domElement
    );
  }
}
