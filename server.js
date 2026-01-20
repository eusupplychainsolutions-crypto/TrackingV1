// server.js (CommonJS) - 覆盖版：个人 OneDrive 用 /content 下载 + 本地解析（不走 /workbook）
const express = require("express");
const cors = require("cors");
const axios = require("axios");
//const ExcelJS = require("exceljs");
const { getGraphToken } = require("./msalClient");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors({ origin: true, credentials: true }));
app.use(express.json());

app.get("/", (req, res) => {
  res.send("Cargo Tracking API is running OK.");
});

// Graph 探针：验证 token + Microsoft Graph 是否可访问
app.get("/api/_health/graph", async (req, res) => {
  try {
    const token = await getGraphToken(["User.Read"]);

    const r = await axios.get("https://graph.microsoft.com/v1.0/me", {
      headers: { Authorization: `Bearer ${token}` },
      timeout: 15000,
    });

    res.json({
      ok: true,
      graph: "reachable",
      user: r.data.userPrincipalName || r.data.mail || r.data.id,
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
 * Excel 探针（个人 OneDrive 推荐）：
 * 1) Graph: 先确认文件存在（metadata）
 * 2) Graph: 用 /content 下载 xlsx（二进制）
 * 3) exceljs: 本地解析，返回 sheetName / 行列数等信息
 *
 * 环境变量：
 * - EXCEL_FILE_NAME: 默认 JobTrackingSample.xlsx
 *   如果在 Documents 下：Documents/JobTrackingSample.xlsx
 */
app.get("/api/_health/excel", async (req, res) => {
  try {
    const token = await getGraphToken(["Files.Read"]);

    const filePath = process.env.EXCEL_FILE_NAME || "JobTrackingSample.xlsx";
    // 注意：Graph 路径里 folder/file 之间用 /，这里不要把整个 path encode 掉
    const encodedPath = filePath
      .split("/")
      .map((p) => encodeURIComponent(p))
      .join("/");

    // 1) 先拿 metadata，确认文件存在 & 拿到大小等信息
    const metaUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedPath}`;
    const metaResp = await axios.get(metaUrl, {
      headers: { Authorization: `Bearer ${token}` },
      timeout: 20000,
    });

    // 2) 下载文件内容（xlsx 二进制）
    const contentUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedPath}:/content`;
    const contentResp = await axios.get(contentUrl, {
      headers: { Authorization: `Bearer ${token}` },
      responseType: "arraybuffer",
      timeout: 30000,
      maxContentLength: 50 * 1024 * 1024, // 50MB
      maxBodyLength: 50 * 1024 * 1024,
    });

    const buffer = Buffer.from(contentResp.data);

    // 3) 解析 xlsx
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    const firstSheet = workbook.worksheets[0];
    const sheetName = firstSheet ? firstSheet.name : null;

    // 取一个轻量信息用于 health check
    const rowCount = firstSheet ? firstSheet.rowCount : 0;
    const colCount = firstSheet ? firstSheet.columnCount : 0;

    // 可选：读取第一行（注意 values[0] 是空占位）
    let firstRow = null;
    if (firstSheet && rowCount >= 1) {
      const values = firstSheet.getRow(1).values;
      // 去掉索引 0 占位，并把 undefined/null 清理一下
      firstRow = Array.isArray(values) ? values.slice(1).map((v) => (v == null ? "" : v)) : null;
    }

    res.json({
      ok: true,
      excel: "download_and_parse_ok",
      file: {
        path: filePath,
        name: metaResp.data.name,
        size: metaResp.data.size,
        lastModifiedDateTime: metaResp.data.lastModifiedDateTime,
      },
      workbook: {
        sheetName,
        rowCount,
        colCount,
        firstRow,
      },
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
