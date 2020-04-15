import React from 'react';
import moment from 'moment';
import config from './Config';
import { getEvents } from './GraphService';

function formatHour(dateTime) {
  return moment.utc(dateTime).local().format('LT');
}

export default class Calendar extends React.Component {

  state = {
    events: []
  };

  async componentDidMount() {
      this.fetchEvents();
      setInterval(this.fetchEvents, 1000*10);
  }

  async fetchEvents() {
    try {
      var accessToken = await window.msal.acquireTokenSilent({
        scopes: config.scopes
      });
      var events = await getEvents(accessToken);
      this.setState({events: events.value});
    }
    catch(err) {
      console.log(err);
    }
  }

  render() {
    return (
      <div className="calendar-container p-3" >
        <h1 className="text-dark">Calendar - {moment().format('MMM D')}</h1>
        <div>
          {this.state.events.map(
            function(event){
              return(
                <div key={event.id} className="event d-flex justify-content-between p-3 mb-3 shadow-lg rounded-sm card-bg">
                  <div>
                    <h4 className="text-red">{event.subject}</h4>
                    <h6 className="text-dark">{event.location ? event.location.displayName : ''}</h6>
                    <div className="font-weight-light">{event.organizer.emailAddress.name}</div>
                  </div>
                  <div>
                    <div>{formatHour(event.start.dateTime)}</div>
                    <div>{formatHour(event.end.dateTime)}</div>
                  </div>
                </div>
              );
            })}
        </div>
      </div>
    );
  }
}