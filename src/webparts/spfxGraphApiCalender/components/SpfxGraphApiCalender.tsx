import * as React from 'react';
import styles from './SpfxGraphApiCalender.module.scss';
import { ISpfxGraphApiCalenderProps } from './ISpfxGraphApiCalenderProps';
import { escape } from '@microsoft/sp-lodash-subset';


import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { ISpFxGraphApiCalendarState } from './ISpFxGraphApiCalendarState';
import { List } from 'office-ui-fabric-react/lib/List';


export default class SpfxGraphApiCalender extends React.Component<ISpfxGraphApiCalenderProps, ISpFxGraphApiCalendarState,{}> {
  
  constructor(props: ISpfxGraphApiCalenderProps) {
    super(props);
    this.state = {
      events: [] 
     };
  }

  public componentDidMount(): void {
    this.props.graphClient
      .api('/me/calendar/events')
      .get((error: any, eventsResponse: any, rawResponse?: any) => {
        const calendarEvents: MicrosoftGraph.Event[] = eventsResponse.value;
        console.log('calendarEvents', calendarEvents);
        this.setState({ events: calendarEvents });
      });
  }

  private _onRenderEventCell(item: MicrosoftGraph.Event, index: number | undefined): JSX.Element {
    return (
       
       <div>
        <h3>{item.subject}</h3>
      </div>

    );
  }

  public render(): React.ReactElement<ISpfxGraphApiCalenderProps> {
    return (
      <List items={this.state.events}
        onRenderCell={this._onRenderEventCell} />
  );
  }
}
