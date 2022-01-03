const redirectUri = 'https://<STORAGEACCOUNTNAME>.z1.web.core.windows.net/';

const msalConfig = {
	auth: {
		clientId: '<CLIENT_ID>',
		authority:
			'https://login.microsoftonline.com/<TENANT_ID>',
		redirectUri,
	},
	cache: {
		cacheLocation: 'sessionStorage',
	},
	system: {
		loggerOptions: {
			loggerCallback(loglevel, message, containsPii) {
				console.log(message);
			},
			piiLoggingEnabled: false,
			logLevel: msal.LogLevel.Verbose,
		},
	},
};

// Add scopes here for ID token to be used at Microsoft identity platform endpoints.
const loginRequest = {
	scopes: ['openid', 'profile', 'User.Read'],
};

// Add scopes here for access token to be used at Microsoft Graph API endpoints.
const tokenRequest = {
	scopes: ['User.Read', 'Group.Read.All'],
};
