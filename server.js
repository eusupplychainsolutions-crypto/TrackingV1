// server.js (CommonJS) — FINAL for A Route (App-only)
// - Uses /users/{upn}/drive (NOT /me)
// - Downloads Excel via /content
// - Optional parse with exceljs if installed

const express = require("express");
const cors = require("cors");
const axios = require("axios");
const { getGraphToken } = require("./msalClient");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors({ origin: true, credentials: true }));
app.use(express.json());

// ===== Config =====
const HEALTH_KEY = process.env.HEALTH_KEY || ""; // optional
const OD_USER_UPN = process.env.OD_USER_UPN || ""; // REQUIRED for app-only
const EXCEL_FILE_NAME = process.env.EXCEL_FILE_NAME || "JobTrackingSample.xlsx";

function requireHealthKey(req, res) {
  if (!HEALTH_KEY) return true; // if not set, allow (for quick testing)
  const key = req.query.key || req.headers["x-health-key"];
  if (key !== HEALTH_KEY) {
    res.status(401).json({ ok: false, message: "Unauthorized: invalid health key" });
    return false;
  }
  return true;
}

function encodeGraphPath(path) {
  return String(path)
    .split("/")
    .filter(Boolean)
    .map((p) => encodeURIComponent(p))
    .join("/");
}

// ===== Routes =====
app.get("/", (req, res) => {
  res.send("Cargo Tracking API is running OK.");
});

/**
 * Graph probe (app-only):
 * - Do NOT call /me (no user context in app-only)
 * - Just validate token + Graph reachable.
 * Requires Sites.Read.All (Application) to succeed.
 */
app.get("/api/_health/graph", async (req, res) => {
  if (!requireHealthKey(req, res)) return;

  try {
    const token = await getGraphToken();

    // Lightweight app-only check (no /me)
    const r = await axios.get("https://graph.microsoft.com/v1.0/sites?search=*", {
      headers: { Authorization: `Bearer ${token}` },
      timeout: 15000,
    });

    res.json({
      ok: true,
      graph: "reachable",
      sitesSampleCount: Array.isArray(r.data?.value) ? r.data.value.length : 0,
    });
  } catch (e) {
    res.status(e.response?.status || 500).json({
      ok: false,
      status: e.response?.status,
      graphError: e.response?.data,
      message: e.message,
      requestUrl: e.config?.url,
    });
  }
});

/**
 * Excel probe (app-only):
 * 1) metadata via /users/{upn}/drive/root:/path
 * 2) download via /content
 * 3) optional parse via exceljs (if installed)
 *
 * REQUIRED env:
 * - OD_USER_UPN (e.g. zhaaojiang@eusupplychainsolutions.onmicrosoft.com)
 *
 * Permissions required (Azure App -> Application permissions):
 * - Files.Read.All  (and usually Sites.Read.All as well)
 * - Admin consent granted
 */
app.get("/api/_health/excel", async (req, res) => {
  if (!requireHealthKey(req, res)) return;

  if (!OD_USER_UPN) {
    return res.status(500).json({
      ok: false,
      message: "Missing env OD_USER_UPN (OneDrive owner UPN).",
    });
  }

  try {
    const token = await getGraphToken();
    const encodedPath = encodeGraphPath(EXCEL_FILE_NAME);

    // 1) metadata
    const metaUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      OD_USER_UPN
    )}/drive/root:/${encodedPath}`;

    const metaResp = await axios.get(metaUrl, {
      headers: { Authorization: `Bearer ${token}` },
      timeout: 20000,
    });

    // 2) download content
    const contentUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      OD_USER_UPN
    )}/drive/root:/${encodedPath}:/content`;

    const contentResp = await axios.get(contentUrl, {
      headers: { Authorization: `Bearer ${token}` },
      responseType: "arraybuffer",
      timeout: 30000,
      maxContentLength: 50 * 1024 * 1024,
      maxBodyLength: 50 * 1024 * 1024,
    });

    const buffer = Buffer.from(contentResp.data);

    // 3) optional parse (exceljs)
    let parse = {
      parsed: false,
      engine: null,
      sheetName: null,
      rowCount: null,
      colCount: null,
      firstRow: null,
      note: null,
    };

    let ExcelJS = null;
    try {
      ExcelJS = require("exceljs");
    } catch (_) {
      // ignore if not installed
    }

    if (ExcelJS) {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      const firstSheet = workbook.worksheets[0];
      if (firstSheet) {
        parse.parsed = true;
        parse.engine = "exceljs";
        parse.sheetName = firstSheet.name;
        parse.rowCount = firstSheet.rowCount;
        parse.colCount = firstSheet.columnCount;

        if (firstSheet.rowCount >= 1) {
          const values = firstSheet.getRow(1).values;
          parse.firstRow = Array.isArray(values)
            ? values.slice(1).map((v) => (v == null ? "" : v))
            : null;
        }
      } else {
        parse.note = "Workbook has no worksheets.";
      }
    } else {
      parse.note = "exceljs not installed; download-only check passed.";
    }

    res.json({
      ok: true,
      excel: "download_ok",
      file: {
        ownerUpn: OD_USER_UPN,
        path: EXCEL_FILE_NAME,
        name: metaResp.data.name,
        size: metaResp.data.size,
        lastModifiedDateTime: metaResp.data.lastModifiedDateTime,
        webUrl: metaResp.data.webUrl,
      },
      parse,
    });
  } catch (e) {
    res.status(e.response?.status || 500).json({
      ok: false,
      status: e.response?.status,
      graphError: e.response?.data,
      message: e.message,
      requestUrl: e.config?.url,
    });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
