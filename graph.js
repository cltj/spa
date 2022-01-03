// Helper function to call Microsoft Graph API endpoint
// using authorization bearer token scheme

async function callMSGraph(endpoint, token) {
	console.log('request made to Graph API at: ' + new Date().toString());

	const bearer = `Bearer ${token}`;

	const headers = {
		ConsistencyLevel: 'eventual',
		Authorization: bearer,
	};

	const options = {
		method: 'GET',
		headers: headers,
		redirect: 'follow',
	};

	/*Group id, displayName and description is returned*/
	try {
		const data = await fetch(endpoint, options).then(res => res.json());

		console.log(data);

		return data;
	} catch (err) {
		console.log(err);
	}
}
