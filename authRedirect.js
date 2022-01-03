//  Create the main myMSALObj instance
// configuration parameters are located at authConfig.js
const myMSALObj = new msal.PublicClientApplication(msalConfig);

let accessToken;
let username = '';

// Redirect: once login is successful and redirects with tokens, call Graph API
myMSALObj
	.handleRedirectPromise()
	.then(handleResponse)
	.catch(err => {
		console.error(err);
	});

async function handleResponse(resp) {
	if (resp !== null) {
		username = resp.account.username;
		await get_PBI();
	} else {
		/**
		 * See here for more info on account retrieval:
		 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
		 */
		const currentAccounts = myMSALObj.getAllAccounts();
		if (currentAccounts === null) {
			return;
		} else if (currentAccounts.length > 1) {
			// Add choose account code here
			console.warn('Multiple accounts detected.');
		} else if (currentAccounts.length === 1) {
			username = currentAccounts[0].username;
			await get_PBI();
		}
	}
}

function signIn() {
	myMSALObj.loginRedirect(loginRequest);
}

function signOut() {
	const logoutRequest = {
		account: myMSALObj.getAccountByUsername(username),
	};
	myMSALObj.logout(logoutRequest);
}

function getTokenRedirect(request) {
	/**
	 * See here for more info on account retrieval:
	 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
	 */
	request.account = myMSALObj.getAccountByUsername(username);
	return myMSALObj.acquireTokenSilent(request).catch(error => {
		console.warn(
			'silent token acquisition fails. acquiring token using redirect'
		);
		if (error instanceof msal.InteractionRequiredAuthError) {
			// fallback to interaction when silent call fails
			return myMSALObj.acquireTokenRedirect(request);
		} else {
			console.warn(error);
		}
	});
}

function updateUI({ value: data }, upn) {
	data.forEach(({ id, displayName, description }) =>
		show_choice(upn, id, displayName, description)
	);
}

function show_choice(upn, groupId, groupDisplayName, groupDescription) {
	htmlStr = `<div class="col-md-4">
	  <div class="card mb-4 box-shadow">
		<div class="card-body">
		  <p class="card-text">${groupDisplayName}</p>
		  <div class="d-flex justify-content-between align-items-center">
			<p class="card-text">${groupDescription}</p>
			<div class="btn-group">
			  <button type="button" onclick="post_choice('${upn}','${groupId}')" class="btn btn-sm btn-outline-secondary">Choose</button>
			</div>
		  </div>
		</div>
	  </div>
	</div>`;

	const groupDiv = document.getElementById('groupdiv');
	groupDiv.innerHTML += htmlStr;
}

async function get_PBI() {
	const res = await getTokenRedirect(tokenRequest);

	const data = await callMSGraph(
		graphConfig.graphGroupEndpoint,
		res.accessToken
	);

	const profile = await getProfile();

	updateUI(data, profile.userPrincipalName);
}

function startInterval(seconds) {
	const timeout = document.getElementById('timeout');
	const result = document.getElementById('result');

	timeout.classList.remove('d-none');
	result.classList.remove('d-none');

	let timeoutSec = seconds;

	const intervalId = setInterval(() => {
		timeout.textContent = `You will be logged out in ${timeoutSec}s`;
		if (timeoutSec === 0) {
			timeout.classList.add('d-none');
			result.classList.add('d-none');
			signOut();
			clearInterval(intervalId);
		}
		timeoutSec--;
	}, 1000);

	groupDiv.innerHTML = '';
}

async function post_choice(upn, choice) {
	const data = { upn, choice };
	console.log(data);

	const requestOptions = {
		method: 'POST',
		redirect: 'follow',
		headers: {
			'Content-Type': 'application/json;charset=utf-8',
		},
		body: JSON.stringify(data),
	};

	const azureUrl =
		'<FINAL_ENDPOINT>';

	/*The endpoint for further processing*/
	const res = await fetch(azureUrl, requestOptions).then(response =>
		response.json()
	);

	const TIMEOUT_SECONDS = 5;

	const groupDiv = document.getElementById('groupdiv');
	const result = document.getElementById('result');

	groupDiv.innerHTML = '';
	result.textContent = res.msg;

	startInterval(TIMEOUT_SECONDS);

	console.log(res);
}

async function getProfile() {
	const res = await new Promise((resolve, reject) =>
		getTokenRedirect(loginRequest)
			.then(response => resolve(response))
			.catch(err => reject(err))
	);

	const profile = await callMSGraph(
		graphConfig.graphMeEndpoint,
		res.accessToken
	);

	return profile;
}
