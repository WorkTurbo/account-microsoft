'use strict';

/**
 * Define the base object namespace. By convention we use the service name
 * in PascalCase (aka UpperCamelCase). Note that this is defined as a package global.
 */

Microsoft = {};

// Request Microsoft credentials for the user
// @param options {optional}
// @param credentialRequestCompleteCallback {Function} Callback function to call on
//   completion. Takes one argument, credentialToken on success, or Error on
//   error.
Microsoft.requestCredential = function (options, credentialRequestCompleteCallback) {
  // support both (options, callback) and (callback).
  if (!credentialRequestCompleteCallback && typeof options === 'function') {
    credentialRequestCompleteCallback = options;
    options = {};
  }

  var config = ServiceConfiguration.configurations.findOne({service: 'microsoft'});
  if (!config) {
    credentialRequestCompleteCallback && credentialRequestCompleteCallback(
      new ServiceConfiguration.ConfigError());
    return;
  }
  var credentialToken = Random.secret();

  var flatScope = config.requestPermissions.map(encodeURIComponent).join('+')

  var loginStyle = OAuth._loginStyle('microsoft', {loginStyle:'popup', ...config}, options);

  var loginUrl =
    //`https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&resource=https%3A%2F%2Foutlook.office.com&client_id=91204085-0ffc-43c9-a52b-3c4cdc1e94e5&redirect_uri=http%3A%2F%2Flocalhost%3A3000%2F_oauth%2Fmicrosoft`
    'https://login.microsoftonline.com/common/oauth2/authorize' +
    '?client_id=' + config.clientId +
    '&response_type=code'+
    //'&scope=' + flatScope +
    '&redirect_uri=' + OAuth._redirectUri('microsoft', {loginStyle:'popup', ...config}) +
    '&state=' + OAuth._stateParam(loginStyle, credentialToken);
  //var loginUrl =
    //'https://login.microsoftonline.com/common/oauth2/v2.0/authorize' +
    //'?client_id=' + config.clientId +
    //'&response_type=code'+
    //'&scope=' + flatScope +
    //'&redirect_uri=' + OAuth._redirectUri('microsoft', {loginStyle:'popup', ...config}) +
    //'&state=' + OAuth._stateParam(loginStyle, credentialToken);

  OAuth.launchLogin({
    loginService: "microsoft",
    loginStyle: loginStyle,
    loginUrl: loginUrl,
    credentialRequestCompleteCallback: credentialRequestCompleteCallback,
    credentialToken: credentialToken,
    popupOptions: {width: 900, height: 450}
  });
};
