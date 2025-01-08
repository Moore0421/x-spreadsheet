const express = require("express");
const fs = require("fs");
const path = require("path");
const app = express();

app.use(express.json());

app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Headers", "Content-Type");
  next();
});

app.use(express.json({ limit: "500mb" }));
app.use(express.urlencoded({ limit: "500mb", extended: true }));
app.use(express.text({ limit: "500mb" }));

// 在顶部添加常量定义数据目录
const DATA_DIR = path.join(__dirname, "data");

// 确保数据目录存在
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR);
}

// 修改 getFileName 方法
function getFileName(id) {
  const fileName = id ? `spreadsheet-data-${id}.json` : "spreadsheet-data.json";
  return path.join(DATA_DIR, fileName);
}

// 保存分片数据
app.post("/api/chunk-upload", (req, res) => {
  const { chunkIndex, totalChunks, id } = req.query;
  const data = JSON.stringify(req.body);
  const tempDir = path.join(__dirname, "temp");

  // 确保临时目录存在
  if (!fs.existsSync(tempDir)) {
    fs.mkdirSync(tempDir);
  }

  const chunkPath = path.join(tempDir, `${id}-${chunkIndex}`);

  try {
    fs.writeFileSync(chunkPath, data);
    res.json({
      success: true,
      chunkIndex,
      message: `Chunk ${chunkIndex} of ${totalChunks} received`,
    });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 合并分片
app.post("/api/merge-chunks", (req, res) => {
  const { id, totalChunks } = req.query;
  const chunksCount = parseInt(totalChunks, 10);
  const tempDir = path.join(__dirname, "temp");
  const finalPath = getFileName(id);
  let rawData = "";

  try {
    console.log("开始合并，总分片数:", chunksCount);

    // 按顺序读取并合并所有分片
    for (let i = 0; i < chunksCount; i++) {
      const chunkPath = path.join(tempDir, `${id}-${i}`);
      const chunkContent = fs.readFileSync(chunkPath, "utf8");
      const chunkData = JSON.parse(chunkContent);

      // 提取分片中的实际数据
      if (chunkData && chunkData.data) {
        rawData += chunkData.data;
      }

      // 删除分片文件
      fs.unlinkSync(chunkPath);
    }

    // 解析完整的JSON字符串
    const finalData = JSON.parse(rawData);
    console.log("合并后数据对象类型:", typeof finalData);

    // 保存为格式化的JSON
    fs.writeFileSync(finalPath, JSON.stringify(finalData, null, 2));
    console.log("最终文件已保存");

    res.json({ success: true, message: "File merged successfully" });
  } catch (error) {
    console.error("合并错误:", error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 获取表格总数
app.get("/api/getSheetCount", (req, res) => {
  const { id } = req.query;
  const fileName = getFileName(id);

  try {
    if (fs.existsSync(fileName)) {
      const data = JSON.parse(fs.readFileSync(fileName, "utf8"));
      res.json({
        success: true,
        count: data.length,
      });
    } else {
      res.json({ success: false });
    }
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 获取单个表格数据
app.get("/api/getSheet", (req, res) => {
  const { id, index } = req.query;
  const fileName = getFileName(id);

  try {
    if (fs.existsSync(fileName)) {
      const data = JSON.parse(fs.readFileSync(fileName, "utf8"));
      if (index >= 0 && index < data.length) {
        res.json({
          success: true,
          sheet: data[index],
        });
      } else {
        res.json({ success: false, message: "Sheet index out of range" });
      }
    } else {
      res.json({});
    }
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.listen(3000, () => {
  console.log("Server running on port 3000");
});
