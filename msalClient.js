// msalClient.js
const { PublicClientApplication } = require("@azure/msal-node");
const fs = require("fs");
const path = require("path");

const cachePath = path.join(__dirname, "msal-cache.json");

const msalConfig = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
  },
  cache: {
    cachePlugin: {
      beforeCacheAccess: async (ctx) => {
        if (fs.existsSync(cachePath)) {
          ctx.tokenCache.deserialize(fs.readFileSync(cachePath, "utf-8"));
        }
      },
      afterCacheAccess: async (ctx) => {
        if (ctx.cacheHasChanged) {
          fs.writeFileSync(cachePath, ctx.tokenCache.serialize());
        }
      },
    },
  },
};

const pca = new PublicClientApplication(msalConfig);

async function getGraphToken() {
  const accounts = await pca.getTokenCache().getAllAccounts();

  if (accounts.length === 0) {
    throw new Error("No cached account found. Run device login first.");
  }

  const result = await pca.acquireTokenSilent({
    account: accounts[0],
    scopes: ["User.Read", "Files.Read"],
  });

  return result.accessToken;
}

module.exports = {
  pca,
  getGraphToken,
};
