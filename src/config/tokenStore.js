const fs = require('fs');
const path = require('path');

// Path to the MSAL token cache file (same as in authConfig)
const TOKEN_CACHE_PATH = path.join(__dirname, 'msalTokenCache.json');

/**
 * Check if the user is authenticated by looking for tokens in the cache file.
 */
function isAuthenticated() {
    try {
        if (fs.existsSync(TOKEN_CACHE_PATH)) {
            const data = fs.readFileSync(TOKEN_CACHE_PATH, 'utf8');
            const cache = JSON.parse(data);
            // Check if there are any refresh tokens in the cache
            return Object.keys(cache).length > 0;
        }
    } catch (error) {
        console.error(`[TokenStore] Failed to read token cache: ${error.message}`);
    }
    return false;
}

/**
 * Get the account information from the token cache.
 */
function getAccount() {
    try {
        if (fs.existsSync(TOKEN_CACHE_PATH)) {
            const data = fs.readFileSync(TOKEN_CACHE_PATH, 'utf8');
            const cache = JSON.parse(data);
            // Look for account data in the cache
            for (const key of Object.keys(cache)) {
                if (key.includes('Account')) {
                    return JSON.parse(cache[key]);
                }
            }
        }
    } catch (error) {
        console.error(`[TokenStore] Failed to read token cache: ${error.message}`);
    }
    return null;
}

/**
 * Set account - not needed with MSAL cache plugin, but kept for API compatibility.
 */
function setAccount(account) {
    // The MSAL cache plugin handles persistence automatically
    console.error('[TokenStore] Account persistence handled by MSAL cache plugin.');
}

module.exports = {
    setAccount,
    getAccount,
    isAuthenticated
};
