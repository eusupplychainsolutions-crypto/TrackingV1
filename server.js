// server.js (CommonJS)
const express = require("express");
const cors = require("cors");
const axios = require("axios");
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

// Excel 探针：按文件名/路径验证 workbook 是否可访问（推荐）
app.get("/api/_health/excel", async (req, res) => {
  try {
    const token = await getGraphToken(["Files.Read"]);

    // 如果文件在根目录：JobTrackingSample.xlsx
    // 如果在 Documents/ 下：Documents/JobTrackingSample.xlsx
    const filePath = process.env.EXCEL_FILE_NAME || "JobTrackingSample.xlsx";

    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(
      filePath
    )}:/workbook`;

    const r = await axios.get(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    res.json({
      ok: true,
      excel: "reachable",
      workbook: r.data.name || filePath,
    });
  } catch (e) {
    res.status(e.response?.status || 500).json({
      ok: false,
      status: e.response?.status,
      graphError: e.response?.data,   // 👈 关键：Graph 的详细原因
      message: e.message,
      requestUrl: e.config?.url,
    });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
