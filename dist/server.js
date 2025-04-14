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
function getFileName(id, isPureData) {
  const fileName = isPureData
    ? `spreadsheet-data-pure-${id}.json`
    : `spreadsheet-data-${id}.json`;
  return path.join(DATA_DIR, fileName);
}

// 保存分片数据
app.post("/api/chunk-upload", (req, res) => {
  const { chunkIndex, totalChunks, id, pure, do: action } = req.query;
  const data = JSON.stringify(req.body);
  const tempDir = path.join(__dirname, "temp");

  // 确保临时目录存在
  if (!fs.existsSync(tempDir)) {
    fs.mkdirSync(tempDir);
  }

  const chunkPath = path.join(tempDir, `${id}-${chunkIndex}`);

  try {
    fs.writeFileSync(chunkPath, data);
    
    // 当pure为true时，表示这是最后一个分片，自动执行合并操作
    if (pure === "true") {
      console.log("接收到最后一个分片，开始自动合并...");
      // 执行合并操作
      const chunksCount = parseInt(totalChunks, 10);
      const isPureData = action === "SavePureData";
      const finalPath = getFileName(id, isPureData);
      let rawData = "";

      try {
        console.log("开始合并，总分片数:", chunksCount);

        // 按顺序读取并合并所有分片
        for (let i = 0; i < chunksCount; i++) {
          const currentChunkPath = path.join(tempDir, `${id}-${i}`);
          const chunkContent = fs.readFileSync(currentChunkPath, "utf8");
          const chunkData = JSON.parse(chunkContent);

          // 提取分片中的实际数据
          if (chunkData && chunkData.data) {
            rawData += chunkData.data;
          }

          // 删除分片文件
          fs.unlinkSync(currentChunkPath);
        }

        // 解析完整的JSON字符串
        const finalData = JSON.parse(rawData);
        console.log("合并后数据对象类型:", typeof finalData);

        // 保存为格式化的JSON
        fs.writeFileSync(finalPath, JSON.stringify(finalData, null, 2));
        console.log("最终文件已保存");

        res.json({ 
          success: true, 
          chunkIndex,
          merged: true,
          message: `Chunk ${chunkIndex} of ${totalChunks} received and all chunks merged successfully` 
        });
      } catch (error) {
        console.error("合并错误:", error);
        res.status(500).json({ success: false, error: error.message });
      }
    } else {
      // 正常返回分片接收成功的消息
      res.json({
        success: true,
        chunkIndex,
        message: `Chunk ${chunkIndex} of ${totalChunks} received`,
      });
    }
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 获取表格总数
app.get("/api/getSheetCount", (req, res) => {
  const { id } = req.query;
  const fileName = getFileName(id, false);
  console.log("fileName", fileName);

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
  const fileName = getFileName(id, false);

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

// 深度比较两个对象是否相同
function isEqual(obj1, obj2) {
  return JSON.stringify(obj1) === JSON.stringify(obj2);
}

// 递归更新对象
function updateObject(fullObj, pureObj) {
  let isModified = false;

  // 特殊处理 rows 对象
  if ("rows" in pureObj && "rows" in fullObj) {
    for (const rowKey in pureObj.rows) {
      if (rowKey in fullObj.rows) {
        // 处理 cells 对象
        if (
          "cells" in pureObj.rows[rowKey] &&
          "cells" in fullObj.rows[rowKey]
        ) {
          for (const cellKey in pureObj.rows[rowKey].cells) {
            if (cellKey in fullObj.rows[rowKey].cells) {
              // 检查并更新 text 字段
              if (
                "text" in pureObj.rows[rowKey].cells[cellKey] &&
                !isEqual(
                  fullObj.rows[rowKey].cells[cellKey].text,
                  pureObj.rows[rowKey].cells[cellKey].text
                )
              ) {
                fullObj.rows[rowKey].cells[cellKey].text =
                  pureObj.rows[rowKey].cells[cellKey].text;
                fullObj.rows[rowKey].cells[cellKey].formattedText =
                  pureObj.rows[rowKey].cells[cellKey].text;
                isModified = true;
              }
            }
          }
        }
      }
    }
  }

  // 处理其他普通字段
  for (const key in pureObj) {
    if (key === "rows") continue; // 跳过已处理的 rows

    if (key in fullObj) {
      if (
        typeof pureObj[key] === "object" &&
        pureObj[key] !== null &&
        typeof fullObj[key] === "object" &&
        fullObj[key] !== null
      ) {
        const { modified, updatedObj } = updateObject(
          fullObj[key],
          pureObj[key]
        );
        if (modified) {
          fullObj[key] = updatedObj;
          isModified = true;
        }
      } else if (!isEqual(fullObj[key], pureObj[key])) {
        fullObj[key] = pureObj[key];
        isModified = true;
      }
    }
  }

  return { modified: isModified, updatedObj: fullObj };
}

// 添加 multer 用于处理文件上传
const multer = require("multer");
const upload = multer({ dest: "uploads/" });

// 同步纯数据到完整文件
app.post("/api/sync-data", upload.single("pureDataFile"), (req, res) => {
  const { id } = req.query;
  const fullFileName = getFileName(id, false);

  try {
    // 检查完整数据文件是否存在
    if (!fs.existsSync(fullFileName)) {
      return res.status(404).json({
        success: false,
        message: "Full data file not found",
      });
    }

    // 读取上传的纯数据文件
    const pureData = JSON.parse(fs.readFileSync(req.file.path, "utf8"));
    // 读取完整数据文件
    const fullData = JSON.parse(fs.readFileSync(fullFileName, "utf8"));

    let modifiedSheets = 0;

    // 遍历纯数据文件中的每个对象
    pureData.forEach((pureSheet) => {
      // 在完整数据中找到对应的对象
      const fullSheet = fullData.find((sheet) => sheet.name === pureSheet.name);

      if (fullSheet) {
        // 递归更新对象
        const { modified, updatedObj } = updateObject(fullSheet, pureSheet);
        if (modified) {
          Object.assign(fullSheet, updatedObj);
          modifiedSheets++;
        }
      }
    });

    // 删除临时上传的文件
    fs.unlinkSync(req.file.path);

    // 只有在有修改时才保存文件
    if (modifiedSheets > 0) {
      fs.writeFileSync(fullFileName, JSON.stringify(fullData, null, 2));
      res.json({
        success: true,
        message: `Successfully synchronized data. Modified ${modifiedSheets} sheets.`,
      });
    } else {
      res.json({
        success: true,
        message: "No changes needed, data already in sync.",
      });
    }
  } catch (error) {
    console.error("同步数据错误:", error);
    // 确保清理临时文件
    if (req.file && req.file.path) {
      fs.unlinkSync(req.file.path);
    }
    res.status(500).json({ success: false, error: error.message });
  }
});

// 获取纯数据
app.get("/api/getPureData", (req, res) => {
  const { id } = req.query;
  const fileName = getFileName(id, true); // 使用纯数据文件名

  try {
    if (fs.existsSync(fileName)) {
      const data = JSON.parse(fs.readFileSync(fileName, "utf8"));
      res.json({
        success: true,
        data: data
      });
    } else {
      res.json({ success: false, message: "纯数据文件不存在" });
    }
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.listen(3000, () => {
  console.log("Server running on port 3000");
});
