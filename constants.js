
	exports.adalConfiguration = {
	  authority: 'https://login.microsoftonline.com/common',

	  redirectUri: 'http://localhost:3000/callback'
	};
	
	exports.subscriptionConfiguration = {
	  changeType: 'Created',
	  notificationUrl: 'https://948b2300.ngrok.io/listen',
	  resource: 'me/mailFolders(\'Inbox\')/messages',
	  clientState: 'cLIENTsTATEfORvALIDATION'
};