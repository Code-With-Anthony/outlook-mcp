const { Client } = require('@microsoft/microsoft-graph-client');
const { pca, scopes } = require('../config/authConfig');
const tokenStore = require('../config/tokenStore');

/**
 * Gets the account from the MSAL cache.
 */
async function getAccountFromCache() {
    try {
        const accounts = await pca.getTokenCache().getAllAccounts();
        console.error(`[GraphClient] Found ${accounts ? accounts.length : 0} accounts in cache`);
        if (accounts && accounts.length > 0) {
            console.error(`[GraphClient] Account: ${JSON.stringify(accounts[0]).substring(0, 200)}`);
            return accounts[0];
        }
    } catch (error) {
        console.error(`[GraphClient] Failed to get account from cache: ${error.message}`);
        console.error(`[GraphClient] Stack: ${error.stack}`);
    }
    return null;
}

/**
 * Gets an access token directly.
 */
async function getAccessToken() {
    try {
        const account = await getAccountFromCache();
        if (!account) {
            throw new Error("No account found in MSAL cache.");
        }

        console.error('[GraphClient] Attempting acquireTokenSilent...');
        const response = await pca.acquireTokenSilent({
            account: account,
            scopes: scopes
        });

        console.error(`[GraphClient] Token acquired: ${response.accessToken ? 'yes' : 'no'}`);
        console.error(`[GraphClient] Token (first 50 chars): ${response.accessToken ? response.accessToken.substring(0, 50) : 'none'}`);
        return response.accessToken;
    } catch (error) {
        console.error("[GraphClient] Failed to acquire token:", error.message);
        throw error;
    }
}

/**
 * Initializes the Microsoft Graph Client with a middleware approach.
 */
const graphClient = Client.initWithMiddleware({
    authProvider: {
        getAccessToken: async () => {
            return await getAccessToken();
        }
    }
});

module.exports = graphClient;
