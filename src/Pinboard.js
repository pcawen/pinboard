import React from 'react';
import { Button } from 'reactstrap';
import config from './Config';
import { getMessages } from './GraphService';
import Calendar from './Calendar';

export default class Pinboard extends React.Component {

  state = {
    messages: [],
    accessToken: ''
  };

  async componentDidMount() {
    this.fetchMessages()
    setInterval(this.fetchMessages.bind(this), 1000*15);
  }

  async fetchMessages() {
    try {
      var accessToken = await window.msal.acquireTokenSilent({
        scopes: config.scopes
      });
      var messages = await getMessages(accessToken);
      this.setState({messages: messages.value});
    }
    catch(err) {
      console.log(err);
    }
  }

  render() {
    if (this.props.isAuthenticated) {
      return (
        <div className="main-container d-flex mt-5">
          <div className="pinboard p-3">
            {this.state.messages.map(
              function(message){
                return(
                  <div className="publication rounded-sm shadow-strong p-3 mb-3 card-bg" key={message.id} dangerouslySetInnerHTML={{ __html: message.body.content }} />
                );
              })}
          </div>
          <Calendar />
        </div>
      );
    } else {
      return <Button color="primary" onClick={this.props.authButtonMethod}>Click here to sign in</Button>
    }
    
  }
}