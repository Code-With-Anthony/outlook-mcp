require('dotenv').config();
const { ConfidentialClientApplication } = require('@azure/msal-node');
const fs = require('fs');
const path = require('path');

// Path to the MSAL token cache file
const TOKEN_CACHE_PATH = path.join(__dirname, 'msalTokenCache.json');

/**
 * Load the token cache from disk
 */
function loadTokenCache() {
    try {
        if (fs.existsSync(TOKEN_CACHE_PATH)) {
            return fs.readFileSync(TOKEN_CACHE_PATH, 'utf8');
        }
    } catch (error) {
        console.error(`[TokenCache] Failed to read token cache: ${error.message}`);
    }
    return '';
}

/**
 * Save the token cache to disk
 */
function saveTokenCache(cacheContents) {
    try {
        fs.writeFileSync(TOKEN_CACHE_PATH, cacheContents);
    } catch (error) {
        console.error(`[TokenCache] Failed to write token cache: ${error.message}`);
    }
}

const msalConfig = {
    auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID || 'common'}`,
        clientSecret: process.env.AZURE_CLIENT_SECRET,
    },
    cache: {
        cachePlugin: {
            beforeCacheAccess: async (cacheContext) => {
                const cacheContents = loadTokenCache();
                cacheContext.tokenCache.deserialize(cacheContents);
            },
            afterCacheAccess: async (cacheContext) => {
                if (cacheContext.cacheHasChanged) {
                    const cacheContents = cacheContext.tokenCache.serialize();
                    saveTokenCache(cacheContents);
                }
            }
        }
    }
};

// Application instance used to run auth flows
const pca = new ConfidentialClientApplication(msalConfig);

module.exports = {
    pca,
    redirectUri: process.env.REDIRECT_URI || "http://localhost:3000/auth/callback",
    // Adding offline_access gives us a Refresh Token specifically
    scopes: ["User.Read", "Mail.Read", "Mail.Send", "Calendars.Read", "Calendars.ReadWrite", "offline_access"]
};
