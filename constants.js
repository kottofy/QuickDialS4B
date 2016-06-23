
exports.adalConfiguration = {
  authority: 'https://login.microsoftonline.com/common',
  clientID : process.env.clientID,
  clientSecret: process.env.clientSecret,
  redirectUri: 'http://quickdial.azurewebsites.net/callback'
};

exports.subscriptionConfiguration = {
  changeType: 'Created',
  notificationUrl: 'https://quickdial.azurewebsites.net/listen',
  resource: 'me/mailFolders(\'Inbox\')/messages',
  clientState: 'cLIENTsTATEfORvALIDATION'
};
