const express = require('express');
const { pca, scopes, redirectUri } = require('../config/authConfig');
const tokenStore = require('../config/tokenStore');

const router = express.Router();

// 1. Kick off the Microsoft OAuth Flow
router.get('/login', async (req, res) => {
    const authCodeUrlParameters = {
        scopes: scopes,
        redirectUri: redirectUri,
    };

    try {
        const response = await pca.getAuthCodeUrl(authCodeUrlParameters);
        res.redirect(response); // Send user to Microsoft Login page
    } catch (error) {
        console.error("Error generating auth url:", error);
        res.status(500).send("Error generating Microsoft login URL.");
    }
});

// 2. Callback from Microsoft with the Authorization Code
router.get('/auth/callback', async (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: scopes,
        redirectUri: redirectUri,
    };

    try {
        // Exchange authorization code for access/refresh tokens
        const response = await pca.acquireTokenByCode(tokenRequest);
        
        // We save the 'account' payload in memory. 
        // Future Graph API requests will pass this account to MSAL to safely request silent tokens!
        tokenStore.setAccount(response.account);

        res.send(`
            <html>
            <head><title>Authentication Complete</title></head>
            <body style="font-family: sans-serif; display: flex; text-align: center; justify-content: center; align-items: center; height: 100vh;">
                <div>
                    <h1 style="color: #0078d4;">✅ Authentication Successful!</h1>
                    <p>The Outlook MCP server is now securely authenticated.</p>
                    <p style="color: gray;">You may close this tab and return to your agent.</p>
                </div>
            </body>
            </html>
        `);
    } catch (error) {
        console.error("Error exchanging auth code:", error);
        res.status(500).send("Failed to exchange authentication code.");
    }
});

module.exports = router;
