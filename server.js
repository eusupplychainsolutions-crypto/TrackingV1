// server.js (CommonJS)
// ClearanceStatus version
// - Uses /users/{upn}/drive (app-only)
// - Downloads Excel via /content
// - Reads ClearanceStatus.xlsx
// - Query by keyValue or customer
// - Returns Customer / KeyValue / DSA1 Code / Updated ETA

const express = require("express");
const cors = require("cors");
const axios = require("axios");
const { getGraphToken } = require("./msalClient");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors({ origin: true, credentials: true }));
app.use(express.json());

// ===== Config =====
const HEALTH_KEY = process.env.HEALTH_KEY || "";
const OD_USER_UPN = process.env.OD_USER_UPN || "";
const EXCEL_FILE_NAME = process.env.EXCEL_FILE_NAME || "ClearanceStatus.xlsx";

function requireHealthKey(req, res) {
  if (!HEALTH_KEY) return true;
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

function ci(v) {
  return String(v ?? "").trim().toLowerCase();
}

function hasText(v) {
  return String(v ?? "").trim() !== "";
}

function normalizeCell(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

function pick(row, keys) {
  for (const k of keys) {
    if (k in row) return row[k];
  }
  return "";
}

app.get("/", (req, res) => {
  res.send("Cargo Tracking API is running OK.");
});

/**
 * Graph health
 */
app.get("/api/_health/graph", async (req, res) => {
  if (!requireHealthKey(req, res)) return;

  try {
    const token = await getGraphToken();

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
 * Excel health
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

    const metaUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      OD_USER_UPN
    )}/drive/root:/${encodedPath}`;

    const metaResp = await axios.get(metaUrl, {
      headers: { Authorization: `Bearer ${token}` },
      timeout: 20000,
    });

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
    } catch (_) {}

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

/**
 * Business API: Clearance
 * GET /api/clearance?key=...&q=...
 * GET /api/clearance?key=...&keyValue=...
 * GET /api/clearance?key=...&customer=...
 *
 * Supports:
 * - q: searches keyValue OR customer
 * - keyValue: searches KeyValue
 * - customer: searches Customer
 *
 * Returns:
 * - Customer
 * - KeyValue
 * - DSA1 Code
 * - Updated ETA
 */
app.get("/api/clearance", async (req, res) => {
  if (!requireHealthKey(req, res)) return;

  try {
    const token = await getGraphToken();

    const ownerUpn = process.env.OD_USER_UPN;
    if (!ownerUpn) {
      return res.status(500).json({ ok: false, message: "Missing env OD_USER_UPN" });
    }

    const filePath = process.env.EXCEL_FILE_NAME || "ClearanceStatus.xlsx";
    const encodedPath = encodeGraphPath(filePath);

    const contentUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      ownerUpn
    )}/drive/root:/${encodedPath}:/content`;

    const contentResp = await axios.get(contentUrl, {
      headers: { Authorization: `Bearer ${token}` },
      responseType: "arraybuffer",
      timeout: 30000,
      maxContentLength: 50 * 1024 * 1024,
      maxBodyLength: 50 * 1024 * 1024,
    });

    const ExcelJS = require("exceljs");
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(Buffer.from(contentResp.data));

    const sheetName = req.query.sheet || workbook.worksheets[0]?.name || "Sheet1";
    const ws = workbook.getWorksheet(sheetName);

    if (!ws) {
      return res.status(404).json({
        ok: false,
        message: `Sheet not found: ${sheetName}`,
      });
    }

    const headerRowIdx = parseInt(req.query.headerRow || "1", 10);

    const headerVals = ws
      .getRow(headerRowIdx)
      .values
      .slice(1)
      .map((v) => (v == null ? "" : String(v).trim()));

    const rows = [];

    for (let r = headerRowIdx + 1; r <= ws.rowCount; r++) {
      const vals = ws.getRow(r).values.slice(1);

      const hasAny = vals.some((v) => v !== null && v !== undefined && String(v).trim() !== "");
      if (!hasAny) continue;

      const obj = {};
      for (let c = 0; c < headerVals.length; c++) {
        const key = headerVals[c] || `col_${c + 1}`;
        obj[key] = vals[c] ?? "";
      }

      // normalize target fields
      const customer = normalizeCell(pick(obj, ["Customer"]));
      const keyValue = normalizeCell(pick(obj, ["KeyValue"]));
      const dsa1Code = normalizeCell(pick(obj, ["DSA1 Code"]));
      const updatedETA = normalizeCell(pick(obj, ["Updated ETA"]));

      // skip junk rows
      if (!hasText(customer) && !hasText(keyValue) && !hasText(dsa1Code) && !hasText(updatedETA)) {
        continue;
      }

      // skip repeated/bad header-like row
      if (ci(customer) === "customer" && ci(keyValue) === "keyvalue") {
        continue;
      }

      // skip weird partial rows that have no useful search values
      if (!hasText(customer) && !hasText(keyValue)) {
        continue;
      }

      rows.push({
        Customer: customer,
        KeyValue: keyValue,
        "DSA1 Code": dsa1Code,
        "Updated ETA": updatedETA,
      });
    }

    const q = req.query.q;
    const qKeyValue = req.query.keyValue;
    const qCustomer = req.query.customer;

    let filtered = rows;

    if (q) {
      filtered = filtered.filter(
        (row) =>
          ci(row["KeyValue"]).includes(ci(q)) ||
          ci(row["Customer"]).includes(ci(q))
      );
    }

    if (qKeyValue) {
      filtered = filtered.filter((row) =>
        ci(row["KeyValue"]).includes(ci(qKeyValue))
      );
    }

    if (qCustomer) {
      filtered = filtered.filter((row) =>
        ci(row["Customer"]).includes(ci(qCustomer))
      );
    }

    const limit = Math.min(parseInt(req.query.limit || "200", 10) || 200, 1000);

    res.json({
      ok: true,
      sheet: sheetName,
      headerRow: headerRowIdx,
      count: filtered.length,
      limit,
      data: filtered.slice(0, limit),
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
