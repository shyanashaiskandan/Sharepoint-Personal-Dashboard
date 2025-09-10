import * as React from 'react';
import styles from './Calendar.module.scss';
import type { ICalendarProps } from './ICalendarProps';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import {Icon} from "@fluentui/react"

export default class Calendar extends React.Component<ICalendarProps, { meetings: any[] }> {
  constructor(props: ICalendarProps) {
    super(props);
    this.state = { meetings: [] };
  }

  public componentDidMount(): void {
    const start = new Date();
    const end = new Date(Date.now() + 24 * 60 * 60 * 1000); // next 24 hours

    this.props.context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3): void => {
        client
          .api('/me/calendarView')
          .query({
            startDateTime: start.toISOString(),
            endDateTime: end.toISOString()
          })
          .select('subject,start,end')
          .top(5)
          .orderby('start/dateTime')
          .get((err, res: any) => {
            if (!err && res?.value) {
              this.setState({ meetings: res.value });
            }
          });
      });
  }

  public render(): React.ReactElement<ICalendarProps> {
    return (
      <section className={styles.calendar}>
        <Icon iconName="Calendar" />
        <h3>Upcoming Meetings (next 24h)</h3>
        {this.state.meetings.length === 0 && <div>No meetings found.</div>}
        <ul>
          {this.state.meetings.map(m => (
            <li key={m.id}>
              <strong>{m.subject || '(No subject)'}</strong><br />
              {new Date(m.start.dateTime).toLocaleString()} – {new Date(m.end.dateTime).toLocaleString()}
            </li>
          ))}
        </ul>
      </section>
    );
  }
}
