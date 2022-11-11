import { AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import { AccountInfo, PublicClientApplication } from "@azure/msal-node";
import { LoginModal, SampleModal } from "LoginModal";
import { App } from "obsidian";
import { TokenCachePlugin } from "TokenCachePlugin";

export default class DeviceCodeAuthProvider implements AuthenticationProvider {
	private accessToken: string;
	private app: App;
	private cachePlugin: TokenCachePlugin;
	private client: PublicClientApplication;
	private scopes: [
		"https://graph.microsoft.com/User.Read",
		"https://graph.microsoft.com/Calendars.Read"
	];
	private modal: LoginModal;

	public onLogout?: () => void;
	public onLogin?: (user: AccountInfo) => void;

	constructor(private clientId: string, app: App) {
		this.app = app;
		this.cachePlugin = new TokenCachePlugin();
		this.modal = new LoginModal(this.app);
		this.initPublicClientApplication();
	}

	public async logout() {
		await this.cachePlugin.deleteFromCache();

		if (this.onLogout) {
			this.onLogout();
		}

	}

	public async getAccessToken(): Promise<string> {
		//check if the pca was initialized (should be done in the constructor)
		if (this.client == null) {
			await this.initPublicClientApplication();
		}

		const accounts = await this.client.getTokenCache().getAllAccounts();

		console.log(accounts);

		if (accounts.length > 0) {
			//we are logged in
			const account = accounts[0];
			await this.getSilently(account);
		} else {
			//we are not logged in
			await this.getByDeviceCode();
		}

		//after we get the token we make sure to close the modal

		this.modal.close();

		const loggedInUser = await this.client.getTokenCache().getAllAccounts();

		if (loggedInUser.length > 0) {
			if (this.onLogin) {
				this.onLogin(loggedInUser[0]);
			}
		}

		return this.accessToken;
	}

	public async getByDeviceCode() {
		const deviceCodeRequest = {
			deviceCodeCallback: (response) => {
				this.modal.setMessage(response.message);
				this.modal.open();
			},
			scopes: this.scopes,
		};

		const response = await this.client.acquireTokenByDeviceCode(
			deviceCodeRequest
		);

		if (response && response.accessToken) {
			this.accessToken = response.accessToken;
		}
	}

	public async initPublicClientApplication() {
		const client = new PublicClientApplication({
			auth: {
				clientId: this.clientId,
				authority:
					"https://login.microsoftonline.com/b36f4c0f-b916-40ae-ad8b-d2f0ea4f8868",
			},
			cache: {
				cachePlugin: this.cachePlugin,
			},
		});
		this.client = client;
	}

	public async getloggedInUser(): Promise<AccountInfo | null> {

		if (this.client != null) {
			const cache = await this.client.getTokenCache();
			if (cache != null) {
				const accounts = await cache.getAllAccounts();
				if (accounts.length > 0) {
					return accounts[0];
				}
			}
		}
		return null;
	}

	public async getSilently(account: AccountInfo) {
		const result = await this.client.acquireTokenSilent({
			scopes: this.scopes,
			account: account,
		});
		if (result) {
			this.accessToken = result.accessToken;
		}
	}
}
