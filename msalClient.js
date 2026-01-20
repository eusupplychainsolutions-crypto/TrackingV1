// msalClient.js (CommonJS) - Render-friendly
const msal = require("@azure/msal-node");

const clientId = process.env.AZURE_CLIENT_ID;
const tenantId = process.env.AZURE_TENANT_ID;
const cacheB64 = process.env.MSAL_CACHE_BASE64;

if (!clientId || !tenantId) throw new Error("Missing AZURE_CLIENT_ID or AZURE_TENANT_ID");
if (!cacheB64) throw new Error("Missing MSAL_CACHE_BASE64");

const cachePlugin = {
  beforeCacheAccess: async (ctx) => {
    const cacheJson = Buffer.from(cacheB64, "base64").toString("utf8");
    ctx.tokenCache.deserialize(cacheJson);
  },
  afterCacheAccess: async () => {
    // Render env 不能回写，这里留空
  },
};

const pca = new msal.PublicClientApplication({
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
  },
  cache: { cachePlugin },
});

async function getGraphToken(scopes = ["User.Read"]) {
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (!accounts.length) {
    throw new Error("No cached account found. MSAL_CACHE_BASE64 not loaded correctly.");
  }
  const result = await pca.acquireTokenSilent({
    account: accounts[0],
    scopes,
  });
  return result.accessToken;
}

module.exports = { pca, getGraphToken };
