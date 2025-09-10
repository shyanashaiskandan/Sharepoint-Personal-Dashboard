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
    const end = new Date(Date.now() + 24 * 60 * 60 * 1000);

    this.props.context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3): void => {
        client
          .api('/me/calendarView')
          .header("Prefer", 'outlook.timezone="America/Toronto"')
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
        <div className={styles.header}>
          <h3 className={styles.headerTitle}>
            <Icon iconName="Calendar" className={styles.calendarIcon} />
            UPCOMING MEETINGS
          </h3>
        </div>
        {this.state.meetings.length === 0 && <div>No meetings found.</div>}
          {this.state.meetings.map(m => (
            <div key={m.id} className={styles.event}>
              <div>
              <strong>{m.subject || '(No subject)'}</strong>
              </div>
              <div>
                {new Date(String(m.start.dateTime)).toLocaleDateString("en-US", {
                  weekday: "short", 
                  month: "short",   
                  day: "numeric",   
                  year: "numeric"   
                })} 
              </div>
              <div>
                {new Date(String(m.start.dateTime)).toLocaleTimeString("en-US", {
                  hour: "numeric",
                  minute: "2-digit"
                })}
                –
                {new Date(String(m.end.dateTime)).toLocaleTimeString("en-US", {
                  hour: "numeric",
                  minute: "2-digit"
                })}
                </div>
            </div>
          ))}
      </section>
    );
  }
}
