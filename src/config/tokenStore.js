/**
 * We store the user's account object in memory.
 * By giving this account object to MSAL Node, it can automatically retrieve 
 * the access token from its cache and handle refreshing tokens behind the scenes.
 */
let currentAccount = null;

module.exports = {
    setAccount: (account) => { currentAccount = account; },
    getAccount: () => currentAccount,
    isAuthenticated: () => currentAccount !== null
};
