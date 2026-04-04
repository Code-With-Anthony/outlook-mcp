const { Client } = require('@microsoft/microsoft-graph-client');
const { pca, scopes } = require('../config/authConfig');
const tokenStore = require('../config/tokenStore');
// We require isomorphic-fetch globally for Microsoft Graph Client if using Node < 18, 
// but since we are on Latest LTS (Node 20+), global 'fetch' is natively supported!

/**
 * Initializes the Microsoft Graph Client.
 * It uses a custom Authentication Provider that automatically fetches 
 * our cached account and acquires a fresh Access Token silently.
 */
const graphClient = Client.init({
    authProvider: async (done) => {
        try {
            if (!tokenStore.isAuthenticated()) {
                const err = new Error("Not authenticated. The user has not logged in via OAuth yet.");
                console.error(err.message);
                return done(err, null);
            }
            
            const account = tokenStore.getAccount();
            
            // acquireTokenSilent automatically looks up the refresh token in the MSAL cache
            // and gets a new access token without requiring user interaction!
            const response = await pca.acquireTokenSilent({
                account: account,
                scopes: scopes
            });
            
            // Return the access token to the Graph Client
            done(null, response.accessToken);
        } catch (error) {
            console.error("Failed to acquire token silently:", error);
            // If this fails, the refresh token might be expired and user must login again.
            done(error, null);
        }
    }
});

module.exports = graphClient;
