// msalClient.js  (CommonJS)

const msal = require("@azure/msal-node");

const clientId = process.env.AZURE_CLIENT_ID;
const tenantId = process.env.AZURE_TENANT_ID;
const cacheB64 = process.env.MSAL_CACHE_BASE64;

if (!clientId || !tenantId) {
  throw new Error("Missing AZURE_CLIENT_ID or AZURE_TENANT_ID");
}
if (!cacheB64) {
  throw new Error("Missing MSAL_CACHE_BASE64");
}

// 用 cachePlugin 把 Render env 里的 token cache 喂给 MSAL
const cachePlugin = {
  beforeCacheAccess: async (cacheContext) => {
    const cacheJson = Buffer.from(cacheB64, "base64").toString("utf8");
    cacheContext.tokenCache.deserialize(cacheJson);
  },
  afterCacheAccess: async () => {
    // Render env 不能回写，所以这里不做 serialize 回写
  },
};

const pca = new msal.PublicClientApplication({
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
  },
  cache: { cachePlugin },
});

module.exports = { pca };

