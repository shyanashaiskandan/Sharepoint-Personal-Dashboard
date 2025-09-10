import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import Banner from './components/Banner';
import { MSGraphClientV3 } from '@microsoft/sp-http';

type Meeting = {
  id: string;
  subject?: string;
  start: { dateTime: string; timeZone?: string };
  end: { dateTime: string; timeZone?: string };
  webLink?: string;
  isCancelled?: boolean;
};

export default class MeetingNotificationApplicationCustomizer
  extends BaseApplicationCustomizer<{}> {

  private _topPlaceholder?: PlaceholderContent;
  private _pollHandle?: number;
  private _currentShownId?: string;

  public async onInit(): Promise<void> {
    this._renderTopPlaceholderShell();
    this._startPolling();
    return;
  }

  public onDispose(): void {
    if (this._pollHandle) window.clearTimeout(this._pollHandle);
    if (this._topPlaceholder?.domElement) {
      ReactDOM.unmountComponentAtNode(this._topPlaceholder.domElement);
    }
  }

  private _renderTopPlaceholderShell(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: () => {} }
      );
    }
  }

  private _startPolling() {
    const tick = async () => {
      try {
        const meeting = await this._getMeetingStartingSoon();
        if (meeting && !this._isDismissed(meeting.id) && this._currentShownId !== meeting.id) {
          this._showBanner(meeting);
        } else if (!meeting) {
          this._hideBannerIfAny();
        }
      } catch {

      }
      finally {
        this._pollHandle = window.setTimeout(tick, 30000);
      }
    };
    tick();
  }

  private async _getMeetingStartingSoon(): Promise<Meeting | undefined> {
    const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

    const now = new Date();
    const end = new Date(now.getTime() + 60 * 60 * 1000);

    const res = await client.api('/me/calendarView')
      .header('Prefer', 'outlook.timezone="America/Toronto"')
      .query({
        startDateTime: now.toISOString(),
        endDateTime: end.toISOString()
      })
      .select('id,subject,start,end,webLink,isCancelled')
      .orderby('start/dateTime')
      .top(10)
      .get();

    const items: Meeting[] = res?.value ?? [];
    const fiveMin = 5 * 60 * 1000;
    const nowMs = Date.now();

    let eligible: Meeting | undefined = undefined;
    for (const evt of items) {
      if (evt.isCancelled) continue;
      const startMs = new Date(evt.start.dateTime).getTime();
      const diff = startMs - nowMs;
      if (diff >= 0 && diff <= fiveMin) {
        eligible = evt;
        break;
      }
    }

    return eligible;
  }

  private _showBanner(meeting: Meeting) {
    if (!this._topPlaceholder?.domElement) return;

    const start = new Date(meeting.start.dateTime);
    const startStr = start.toLocaleTimeString('en-CA', { hour: '2-digit', minute: '2-digit' });

    const message = `“${meeting.subject || 'No subject'}” starts at ${startStr}`;
    const onDismiss = () => {
      this._currentShownId = meeting.id;
      this._rememberDismiss(meeting.id);
      this._hideBannerIfAny();
    };

    ReactDOM.render(
      React.createElement(Banner, { message, onDismiss }),
      this._topPlaceholder.domElement
    );

    this._currentShownId = meeting.id;
  }

  private _hideBannerIfAny() {
    if (this._topPlaceholder?.domElement) {
      ReactDOM.unmountComponentAtNode(this._topPlaceholder.domElement);
    }
    this._currentShownId = undefined;
  }

  private _rememberDismiss(meetingId: string) {
    try { sessionStorage.setItem(`meetingBannerDismissed:${meetingId}`, '1'); } catch {}
  }

  private _isDismissed(meetingId: string): boolean {
    try { return sessionStorage.getItem(`meetingBannerDismissed:${meetingId}`) === '1'; } catch { return false; }
  }
}
