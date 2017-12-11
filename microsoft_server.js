'use strict';
MicrosoftGraph = require('@microsoft/microsoft-graph-client')

Microsoft = {};

OAuth.registerService('microsoft', 2, null, function(query) {
  var respData = getAccessToken(query);
  var c = MicrosoftGraph.Client.init({
    authProvider: function(done) {
      return done(null, respData.access_token)
    }
  })
  var me = Promise.await(c.api('/me').get())

  return {
    serviceData: {
      id: me.id,
      accessToken: OAuth.sealSecret(respData.access_token),
      email: me.mail,
      provider: respData.provider,
      tokenData: respData,
      name: me.displayName
    },
    options: {
      profile: { email: respData.email_address, name: me.displayName }
    }
  };
});

// http://developer.github.com/v3/#user-agent-required
var userAgent = 'Meteor';
if (Meteor.release) userAgent += '/' + Meteor.release;

var getIdentity = function(namespace, accessToken) {
  try {
    return HTTP.get('https://graph.microsoft.com/v1.0/me', {
      headers: {
        Accept: 'application/json',
        'User-Agent': userAgent
      }, // http://developer.github.com/v3/#user-agent-required
      auth: accessToken + ':'
    }).data;
  } catch (err) {
    throw _.extend(
      new Error('Failed to fetch identity from Microsoft. ' + err.message),
      { response: err.response }
    );
  }
};

var getAccessToken = function(query) {
  var config = ServiceConfiguration.configurations.findOne({
    service: 'microsoft'
  });
  if (!config) throw new ServiceConfiguration.ConfigError();

  var response;
  try {
    response = HTTP.post(
      'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      {
        headers: {
          Accept: 'application/json',
          'User-Agent': userAgent
        },
        params: {
          code: query.code,
          grant_type: 'authorization_code',
          client_id: config.clientId,
          client_secret: config.secret,
          redirect_uri: config.redirect_uri
        }
      }
    );
  } catch (err) {
    throw _.extend(
      new Error(
        'Failed to complete OAuth handshake with Microsoft. ' + err.message
      ),
      { response: err.response }
    );
  }
  if (response.data.error) {
    // if the http response was a json object with an error attribute
    throw new Error(
      'Failed to complete OAuth handshake with Microsoft. ' +
        response.data.reason
    );
  } else {
    return response.data;
  }
};

Microsoft.retrieveCredential = function(credentialToken, credentialSecret) {
  return OAuth.retrieveCredential(credentialToken, credentialSecret);
};
