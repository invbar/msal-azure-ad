import * as msal from '@azure/msal-browser'
let msalApp

export default {
  async configure(clientId) {
    if (msalApp) {
      return
    }
    if (!clientId) {
      return
    }

    const config = {
      auth: {
        clientId: clientId,
        redirectUri: window.location.origin,
        authority: 'https://login.microsoftonline.com/tenantid/'
      },
      cache: {
        cacheLocation: 'localStorage'
      }
    }
    console.log('### Azure AD sign-in: enabled\n', config)

    // Create our shared/static MSAL app object
    msalApp = new msal.PublicClientApplication(config)
  },

  //
  // Return the configured client id
  //
  clientId() {
    if (!msalApp) {
      return null
    }

    return msalApp.config.auth.clientId
  },

  //
  // Login a user with a popup
  //
  async login(scopes = ['user.read', 'openid', 'profile', 'email' ]) {
    if (!msalApp) {
      return
    }

    //const LOGIN_SCOPES = ['user.read', 'openid', 'profile', 'email']
    await msalApp.loginPopup({
      scopes,
      prompt: 'select_account'
    })
  },

  //
  // Logout any stored user
  //
  logout() {
    if (!msalApp) {
      return
    }

    msalApp.logoutPopup()
  },

  //
  // Call to get user, probably cached and stored locally by MSAL
  //
  user() {
    if (!msalApp) {
      return null
    }

    const currentAccounts = msalApp.getAllAccounts()
    if (!currentAccounts || currentAccounts.length === 0) {
      // No user signed in
      return null
    } else if (currentAccounts.length > 1) {
      return currentAccounts[0]
    } else {
      return currentAccounts[0]
    }
  },

  //
  // Call through to acquireTokenSilent or acquireTokenPopup
  //
  async acquireToken() {
    if (!msalApp) {
      return null
    }

    // Set scopes for token request
    // const accessTokenRequest = {
    //   scopes,
    //   account: this.user()
    // }

    let tokenResp
    try {
    //   var request = {
    //     scopes: [ "email", "openid" ,"profile", "User.Read", "User.ReadBasic.All"]
    // };
    
      // 1. Try to acquire token silently
      tokenResp = await msalApp.acquireTokenSilent({
        scopes: ["email", "profile", "User.Read", "User.ReadBasic.All"]
      })
     console.log(tokenResp.accessToken)
      console.log('### MSAL acquireTokenSilent was successful')
    } catch (err) {
      // 2. Silent process might have failed so try via popup
      tokenResp = await msalApp.acquireTokenPopup({
        scopes: [ "openid", "email" ,"profile", "User.Read", "User.ReadBasic.All"]
      })
      console.log(tokenResp.accessToken)
      console.log('### MSAL acquireTokenPopup was successful')
    }

    // Just in case check, probably never triggers
    if (!tokenResp.accessToken) {
      throw new Error("### accessToken not found in response, that's bad")
    }

    return tokenResp.accessToken
  },

  //
  // Clear any stored/cached user
  //
  clearLocal() {
    if (msalApp) {
      for (let entry of Object.entries(localStorage)) {
        let key = entry[0]
        if (key.includes('login.windows')) {
          localStorage.removeItem(key)
        }
      }
    }
  },

  //
  // Check if we have been setup & configured
  //
  isConfigured() {
    return msalApp != null
  }
}
