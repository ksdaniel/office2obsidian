import { TokenCacheContext, ICachePlugin } from "@azure/msal-common";
import { safeStorage } from "@electron/remote";

export class TokenCachePlugin implements ICachePlugin {
	public displayName = "TokenCachePlugin";
	private STORAGE_KEY = "msal_token_cache_";
	public acquired: boolean;

	public async beforeCacheAccess(
		cacheContext: TokenCacheContext
	): Promise<void> {
		console.log("beforeCacheAccess");

		const encryptedCache: string | null = localStorage.getItem(
			this.STORAGE_KEY
		);
		const cache =
			encryptedCache !== null
				? safeStorage.decryptString(
						Buffer.from(encryptedCache, "latin1")
				  )
				: "";

		cacheContext.tokenCache.deserialize(cache);
	}

	public async afterCacheAccess(
		cacheContext: TokenCacheContext
	): Promise<void> {
		console.log("afterCacheAccess");

		if (cacheContext.cacheHasChanged) {
			const serializedAccounts = cacheContext.tokenCache.serialize();

			localStorage.setItem(
				this.STORAGE_KEY,
				safeStorage.encryptString(serializedAccounts).toString("latin1")
			);
		}
	}

	public async deleteFromCache(): Promise<void> {
		await localStorage.removeItem(this.STORAGE_KEY);
	}

	public async cacheExists(): Promise<boolean> {
		return (await localStorage.getItem(this.STORAGE_KEY)) !== null;
	}

	constructor() {
		this.acquired = false;
	}
}
