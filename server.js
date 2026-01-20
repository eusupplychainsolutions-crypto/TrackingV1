// server.js
const express = require("express");
const axios = require("axios");
const { getGraphToken } = require("./msalClient");

const app = express();
const PORT = process.env.PORT || 3000;

/**
 * 基础健康检查
 */
app.get("/", (req, res) => {
  res.send("Cargo Tracking API is running OK.");
});

/**
 * Graph 探针接口（核心）
 * 目标：能不能通过 Graph 拿到 Excel workbook 信息
 */
app.get("/api/_health/graph", async (req, res) => {
  try {
    const token = await getGraphToken();

    //  这里用最“轻”的 Graph 调用，不读表，只确认 workbook 存在
    const graphRes = await axios.get(
      `https://graph.microsoft.com/v1.0/me/drive/root:/JobTrackingSample.xlsx:/workbook`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
        },
      }
    );

    res.json({
      ok: true,
      graph: "reachable",
      workbook: graphRes.data.name,
    });
  } catch (err) {
    res.status(500).json({
      ok: false,
      error: err.message,
    });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
