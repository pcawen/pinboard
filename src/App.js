import React, { Component } from 'react';
import { UserAgentApplication } from 'msal';
import './App.css';
import NavBar from './NavBar';
import config from './Config';
import { getUserDetails } from './GraphService';
import 'bootstrap/dist/css/bootstrap.css';
import Pinboard from './Pinboard';

class App extends Component {

  state = {
    isAuthenticated: null,
    user: {},
    error: null
  };

  userAgentApplication = new UserAgentApplication({
    auth: {
        clientId: config.appId,
        redirectUri: config.redirectUri
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true
    }
  });

  componentDidMount() {
    var user = this.userAgentApplication.getAccount();
    if (user) {
      // Enhance user object with data from Graph
      this.getUserProfile();
    }
    this.setState( { isAuthenticated: (user !== null) } )
  }

  render() {
    return (
      <div className="main-container">
        <NavBar
          isAuthenticated={this.state.isAuthenticated}
          authButtonMethod={this.state.isAuthenticated ? this.logout.bind(this) : this.login.bind(this)}
          user={this.state.user}/>
        <div>
          <Pinboard {...this.props}
            isAuthenticated={this.state.isAuthenticated}
            authButtonMethod={this.login.bind(this)} />
        </div>
      </div>
    );
  }

  async login() {
    try {
      await this.userAgentApplication.loginPopup(
        {
          scopes: config.scopes,
          prompt: "select_account"
      });
      await this.getUserProfile();
    }
    catch(err) {
      var error = {};

      if (typeof(err) === 'string') {
        var errParts = err.split('|');
        error = errParts.length > 1 ?
          { message: errParts[1], debug: errParts[0] } :
          { message: err };
      } else {
        error = {
          message: err.message,
          debug: JSON.stringify(err)
        };
      }

      this.setState({
        isAuthenticated: false,
        user: {},
        error: error
      });
    }
  }

  logout() {
    this.userAgentApplication.logout();
  }

  async getUserProfile() {
    try {
      // Get the access token silently
      // If the cache contains a non-expired token, this function
      // will just return the cached token. Otherwise, it will
      // make a request to the Azure OAuth endpoint to get a token

      var accessToken = await this.userAgentApplication.acquireTokenSilent({
        scopes: config.scopes
      });

      if (accessToken) {
        // Get the user's profile from Graph
        var user = await getUserDetails(accessToken);
        this.setState({
          isAuthenticated: true,
          user: {
            displayName: user.displayName,
            email: user.mail || user.userPrincipalName
          },
          error: null
        });
      }
    }
    catch(err) {
      var error = {};
      if (typeof(err) === 'string') {
        var errParts = err.split('|');
        error = errParts.length > 1 ?
          { message: errParts[1], debug: errParts[0] } :
          { message: err };
      } else {
        error = {
          message: err.message,
          debug: JSON.stringify(err)
        };
      }

      this.setState({
        isAuthenticated: false,
        user: {},
        error: error
      });
    }
  }
}

export default App;

