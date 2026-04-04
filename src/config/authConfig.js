require('dotenv').config();
const { ConfidentialClientApplication } = require('@azure/msal-node');

const msalConfig = {
    auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID || 'common'}`,
        clientSecret: process.env.AZURE_CLIENT_SECRET,
    }
};

// Application instance used to run auth flows
const pca = new ConfidentialClientApplication(msalConfig);

module.exports = {
    pca,
    redirectUri: process.env.REDIRECT_URI || "http://localhost:3000/auth/callback",
    // Adding offline_access gives us a Refresh Token specifically
    scopes: ["User.Read", "Mail.Read", "Mail.Send", "Calendars.Read", "offline_access"]
};
