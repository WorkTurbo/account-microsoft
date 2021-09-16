Package.describe({
    name: 'workturbo:microsoft',
    version: '0.0.1',
    summary: 'Login using microsoft'
})

Npm.depends({
  '@microsoft/microsoft-graph-client': '1.0.0'
})
Package.onUse(function(api) {
    api.versionsFrom(['1.2.1', '2.3.1']);
    api.use('ecmascript');
    api.use('oauth2', ['client', 'server']);
    api.use('oauth', ['client', 'server']);
    api.use('http', ['server']);
    api.use('accounts-base', ['client', 'server']);
    // Export Accounts (etc) to packages using this one.
    api.imply('accounts-base', ['client', 'server']);
    api.use('accounts-oauth', ['client', 'server']);
    api.use(['underscore', 'service-configuration'], ['client', 'server']);
    api.use(['random', 'templating'], 'client');

    api.addFiles('account_microsoft.js')
    api.addFiles('microsoft_server.js', 'server')
    api.addFiles(['microsoft_client.js', 'microsoft_login.css'], 'client')
    api.export('Microsoft')
})

