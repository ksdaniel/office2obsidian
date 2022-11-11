import { Configuration, PublicClientApplication } from "@azure/msal-node";

export const getTokenDeviceCode = async () => {
	const clientConfig: Configuration = {
		auth: {
			clientId: "e372a8e0-5b7b-4e04-b89e-0a9c37535eb9",
			authority:
				"https://login.microsoftonline.com/b36f4c0f-b916-40ae-ad8b-d2f0ea4f8868",
		},
	};

	const pca = new PublicClientApplication(clientConfig);

	await pca.acquireTokenByDeviceCode({
		scopes: ["https://graph.microsoft.com/User.Read"],
		deviceCodeCallback: (response) => {
			console.log(response.message);
		},
	});

	return pca;
};
