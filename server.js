// server.js (CommonJS)
const express = require("express");
const cors = require("cors");
const axios = require("axios");
const { getGraphToken } = require("./msalClient");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors({ origin: true, credentials: true }));
app.use(express.json());

/**
 * 基础健康检查
 */
app.get("/", (req, res) => {
  res.send("Cargo Tracking API is running OK.");
});

/**
 * Graph 探针：验证 token + Microsoft Graph 是否可访问
 * 成功返回当前登录用户信息（不会泄露 token）
 */
app.get("/api/_health/graph", async (req, res) => {
  try {
    const token = await getGraphToken(["User.Read"]);

    const r = await axios.get("https://graph.microsoft.com/v1.0/me", {
      headers: { Authorization: `Bearer ${token}` },
    });

    res.json({
      ok: true,
      graph: "reachable",
      user: r.data.userPrincipalName || r.data.mail || r.data.id,
    });
  } catch (e) {
    res.status(500).json({
      ok: false,
      error: e.message,
    });
  }
});

/**
 * Excel 探针：验证能否访问某个 OneDrive Excel Workbook
 * 你可以用 FILE_NAME 或 ITEM_ID 二选一：
 * - 推荐：EXCEL_FILE_NAME=JobTrackingSample.xlsx
 * - 或：EXCEL_ITEM_ID=xxxxx!sxxxx
 */
app.get("/api/_health/excel", async (req, res) => {
  try {
    const token = await getGraphToken(["User.Read", "Files.Read"]);

    const itemId = process.env.EXCEL_ITEM_ID; // 可选
    const fileName = process.env.EXCEL_FILE_NAME || "JobTrackingSample.xlsx"; // 可选

    let url;

    if (itemId) {
      // 用 itemId 定位
      url = `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(
        itemId
      )}/workbook`;
    } else {
      // 用文件名定位（文件必须在 OneDrive 根目录或你自己改路径）
      url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(
        fileName
      )}:/workbook`;
    }

    const r = await axios.get(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    res.json({
      ok: true,
      excel: "reachable",
      workbook: r.data.name || fileName,
    });
  } catch (e) {
    res.status(500).json({
      ok: false,
      error: e.message,
    });
  }
});
app.get("/api/_health/excel", async (req, res) => {
  try {
    const token = await getGraphToken(["User.Read", "Files.Read"]);

    const itemId = process.env.EXCEL_ITEM_ID;
    if (!itemId) {
      return res.status(400).json({
        ok: false,
        error: "Missing EXCEL_ITEM_ID in Render environment variables",
      });
    }

    // 只验证 workbook 是否可访问
    const url = `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(
      itemId
    )}/workbook`;

    const r = await axios.get(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    res.json({
      ok: true,
      excel: "reachable",
      workbook: r.data.name || "workbook",
    });
  } catch (e) {
    res.status(500).json({
      ok: false,
      error: e.message,
    });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
