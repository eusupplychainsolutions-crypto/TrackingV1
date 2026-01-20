import { pca } from "../auth/msalClient.js";

router.get("/api/_health/graph", async (req, res) => {
  try {
    const result = await pca.acquireTokenSilent({
      scopes: ["User.Read"],
      account: pca.getTokenCache().getAllAccounts()[0],
    });

    res.json({
      ok: true,
      graph: "reachable",
      expiresOn: result.expiresOn,
    });
  } catch (err) {
    res.status(500).json({
      ok: false,
      error: err.message,
    });
  }
});
