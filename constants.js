
	exports.adalConfiguration = {
	  authority: 'https://login.microsoftonline.com/common',
   
	 redirectUri: 'http://localhost:3000/callback'
	};
	
	exports.subscriptionConfiguration = {
	  changeType: 'Created',
	  notificationUrl: 'https://9a0b3c14.ngrok.io/listen',
	  resource: 'me/mailFolders(\'Inbox\')/messages',
	  clientState: 'cLIENTsTATEfORvALIDATION'
};