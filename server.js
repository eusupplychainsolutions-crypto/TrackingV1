// server.js (CommonJS)
// ClearanceStatus version
// Stable DD/MM/YYYY handling for Excel ETA values

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
    res.status(401).json({
      ok: false,
      message: "Unauthorized: invalid health key"
    });
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

function formatDateDMY(dd, mm, yyyy) {
  const day = String(dd).padStart(2, "0");
  const month = String(mm).padStart(2, "0");
  return `${day}/${month}/${yyyy}`;
}

function normalizeExcelValue(value, cell) {
  if (value === null || value === undefined) return "";

  const text =
    cell && typeof cell.text === "string"
      ? cell.text.replace(/\u00A0/g, " ").trim()
      : "";

  if (text) return text;

  if (
    Object.prototype.toString.call(value) === "[object Date]" &&
    !isNaN(value.getTime())
  ) {
    return formatDateDMY(
      value.getUTCDate(),
      value.getUTCMonth() + 1,
      value.getUTCFullYear()
    );
  }

  return String(value).trim();
}

// ===== SPECIAL ETA HANDLER =====
function normalizeUpdatedETACell(cell) {
  if (!cell) return "";

  const rawText = String(cell.text || "")
    .replace(/\u00A0/g, " ")
    .trim();

  // Always treat slash dates as DD/MM/YYYY
  let m = rawText.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

  if (m) {
    const dd = Number(m[1]);
    const mm = Number(m[2]);
    const yyyy = Number(m[3]);

    return formatDateDMY(dd, mm, yyyy);
  }

  // DD/MM/YY
  m = rawText.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);

  if (m) {
    const dd = Number(m[1]);
    const mm = Number(m[2]);
    const yy = Number(m[3]);

    const yyyy = yy >= 70 ? 1900 + yy : 2000 + yy;

    return formatDateDMY(dd, mm, yyyy);
  }

  const value = cell.value;

  // Real Excel Date object
  if (
    Object.prototype.toString.call(value) === "[object Date]" &&
    !isNaN(value.getTime())
  ) {
    return formatDateDMY(
      value.getUTCDate(),
      value.getUTCMonth() + 1,
      value.getUTCFullYear()
    );
  }

  const cleaned = String(value ?? rawText ?? "")
    .replace(/\u00A0/g, " ")
    .trim();

  if (!cleaned) return "";

  // ISO
  m = cleaned.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/);

  if (m) {
    return formatDateDMY(
      Number(m[3]),
      Number(m[2]),
      Number(m[1])
    );
  }

  // English Date string
  if (/^[A-Za-z]{3}\s[A-Za-z]{3}\s\d{1,2}\s\d{4}/.test(cleaned)) {
    const parsed = new Date(cleaned);

    if (!isNaN(parsed.getTime())) {
      return formatDateDMY(
        parsed.getUTCDate(),
        parsed.getUTCMonth() + 1,
        parsed.getUTCFullYear()
      );
    }
  }

  return cleaned;
}

app.get("/", (req, res) => {
  res.send("Cargo Tracking API is running OK.");
});

/**
 * Business API: Clearance
 */
app.get("/api/clearance", async (req, res) => {
  if (!requireHealthKey(req, res)) return;

  try {
    const token = await getGraphToken();

    const ownerUpn = process.env.OD_USER_UPN;

    if (!ownerUpn) {
      return res.status(500).json({
        ok: false,
        message: "Missing env OD_USER_UPN"
      });
    }

    const filePath =
      process.env.EXCEL_FILE_NAME || "ClearanceStatus.xlsx";

    const encodedPath = encodeGraphPath(filePath);

    const contentUrl =
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
        ownerUpn
      )}/drive/root:/${encodedPath}:/content`;

    const contentResp = await axios.get(contentUrl, {
      headers: {
        Authorization: `Bearer ${token}`
      },
      responseType: "arraybuffer",
      timeout: 30000,
      maxContentLength: 50 * 1024 * 1024,
      maxBodyLength: 50 * 1024 * 1024
    });

    const ExcelJS = require("exceljs");

    const workbook = new ExcelJS.Workbook();

    await workbook.xlsx.load(Buffer.from(contentResp.data));

    const sheetName =
      req.query.sheet ||
      workbook.worksheets[0]?.name ||
      "Sheet1";

    const ws = workbook.getWorksheet(sheetName);

    if (!ws) {
      return res.status(404).json({
        ok: false,
        message: `Sheet not found: ${sheetName}`
      });
    }

    const headerRowIdx = parseInt(
      req.query.headerRow || "1",
      10
    );

    const headerVals = ws
      .getRow(headerRowIdx)
      .values
      .slice(1)
      .map((v) =>
        v == null ? "" : String(v).trim()
      );

    const rows = [];

    for (let r = headerRowIdx + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);

      const rawVals = row.values.slice(1);

      const hasAny = rawVals.some(
        (v) =>
          v !== null &&
          v !== undefined &&
          String(v).trim() !== ""
      );

      if (!hasAny) continue;

      const obj = {};

      for (let c = 0; c < headerVals.length; c++) {
        const key = headerVals[c] || `col_${c + 1}`;

        const cell = row.getCell(c + 1);

        obj[key] = normalizeExcelValue(
          cell.value,
          cell
        );
      }

      const customer = normalizeCell(
        pick(obj, ["Customer"])
      );

      const keyValue = normalizeCell(
        pick(obj, ["KeyValue"])
      );

      const dsa1Code = normalizeCell(
        pick(obj, ["DSA1 Code"])
      );

      // ===== SPECIAL ETA =====
      const updatedEtaColIndex =
        headerVals.findIndex(
          (h) => h === "Updated ETA"
        ) + 1;

      const updatedEtaCell =
        updatedEtaColIndex > 0
          ? row.getCell(updatedEtaColIndex)
          : null;

      const updatedETA =
        normalizeUpdatedETACell(updatedEtaCell);

      if (
        !hasText(customer) &&
        !hasText(keyValue) &&
        !hasText(dsa1Code) &&
        !hasText(updatedETA)
      ) {
        continue;
      }

      if (
        ci(customer) === "customer" &&
        ci(keyValue) === "keyvalue"
      ) {
        continue;
      }

      if (
        !hasText(customer) &&
        !hasText(keyValue)
      ) {
        continue;
      }

      rows.push({
        Customer: customer,
        KeyValue: keyValue,
        "DSA1 Code": dsa1Code,
        "Updated ETA": updatedETA
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

    const limit = Math.min(
      parseInt(req.query.limit || "200", 10) || 200,
      1000
    );

    res.json({
      ok: true,
      sheet: sheetName,
      headerRow: headerRowIdx,
      count: filtered.length,
      limit,
      data: filtered.slice(0, limit)
    });
  } catch (e) {
    res.status(e.response?.status || 500).json({
      ok: false,
      status: e.response?.status,
      graphError: e.response?.data,
      message: e.message,
      requestUrl: e.config?.url
    });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
