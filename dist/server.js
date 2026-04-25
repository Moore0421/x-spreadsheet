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

// 根据业务规则处理纯数据
function processPureData(data) {
  if (!Array.isArray(data)) return data;

  // 找到相关的表格
  const coverSheet = data.find(sheet => sheet.name && sheet.name.includes("封面代码"));
  const singleSheet = data.find(sheet => sheet.name && sheet.name.includes("资产情况单户表"));
  const summarySheet = data.find(sheet => sheet.name && sheet.name.includes("资产情况汇总表"));
  const distSheet = data.find(sheet => sheet.name && sheet.name.includes("地方资产分布情况表"));

  if (!coverSheet || !singleSheet) {
    return data;
  }

  // 检查封面代码 pxl 11, F4 是否为行政单位或事业单位或民间非盈利组织
  const row11 = coverSheet.rows.find(row => row.pxl === 11);
  const isAdministrativeUnit = row11 && row11.F4 === "行政单位"; 
  const isPublicInstitution = row11 && row11.F4 === "事业单位";
  const isNonProfitOrganization = row11 && row11.F4 === "民间非盈利组织";

  // 检查封面代码 pxl 13, F8 (层级)
  const row13 = coverSheet.rows.find(row => row.pxl === 13);
  const level = row13 ? row13.F8 : null;

  if (summarySheet && (isAdministrativeUnit || isPublicInstitution || isNonProfitOrganization)) {
    // 遍历单户表的所有行
    singleSheet.rows.forEach(singleRow => {
      // 查找汇总表中对应 pxl 的行
      let summaryRow = summarySheet.rows.find(row => row.pxl === singleRow.pxl);

      if (!summaryRow) {
        // 如果汇总表中没有对应 pxl 的行，则创建一行
        summaryRow = { pxl: singleRow.pxl };
        summarySheet.rows.push(summaryRow);
      }

      if (isAdministrativeUnit) {
        // 将单户表的 F3, F4 填充到汇总表的 F7, F8
        if (singleRow.F3 !== undefined) {
          summaryRow.F7 = singleRow.F3;
        }
        if (singleRow.F4 !== undefined) {
          summaryRow.F8 = singleRow.F4;
        }
      } else if (isPublicInstitution) {
        // 将单户表的 F3, F4 填充到汇总表的 F11, F12
        if (singleRow.F3 !== undefined) {
          summaryRow.F11 = singleRow.F3;
        }
        if (singleRow.F4 !== undefined) {
          summaryRow.F12 = singleRow.F4;
        }
      } else if (isNonProfitOrganization) {
        // 将单户表的 F7, F8 填充到汇总表的 F15, F16
        if (singleRow.F7 !== undefined) {
          summaryRow.F15 = singleRow.F7;
        }
        if (singleRow.F8 !== undefined) {
          summaryRow.F16 = singleRow.F8;
        }
      }

      const getNum = (val) => {
        if (val === undefined || val === null || val === "——" || val === "") return 0;
        const num = Number(val);
        return isNaN(num) ? 0 : num;
      };

      const f7 = getNum(summaryRow.F7);
      const f11 = getNum(summaryRow.F11);
      const f15 = getNum(summaryRow.F15);
      summaryRow.F3 = f7 + f11 + f15;

      const f8 = getNum(summaryRow.F8);
      const f12 = getNum(summaryRow.F12);
      const f16 = getNum(summaryRow.F16);
      summaryRow.F4 = f8 + f12 + f16;
    });
  }

  // 地方资产分布情况表逻辑
  if (distSheet && level && (isAdministrativeUnit || isPublicInstitution || isNonProfitOrganization)) {
    // 确定目标列
    let targetCol1, targetCol2;
    if (level === "省级") {
      targetCol1 = "F5";
      targetCol2 = "F6";
    } else if (level === "市级") {
      targetCol1 = "F7";
      targetCol2 = "F8";
    } else if (level === "县级" || level === "乡镇级") {
      targetCol1 = "F9";
      targetCol2 = "F10";
    }

    if (targetCol1 && targetCol2) {
      // 确定行偏移: 行政单位单户表(pxl 6-46) 对应 情况表(pxl 10-50)，即偏移为+4
      // 事业单位单户表(pxl 6-46) 对应 情况表(pxl 54-94)，即偏移为+48
      // 民间非盈利组织单户表(pxl 6-46) 对应 情况表(pxl 100-140)，即偏移为+40
      const rowOffset = isAdministrativeUnit ? 4 : (isPublicInstitution ? 48 : (isNonProfitOrganization ? 40 : 0));

      singleSheet.rows.forEach(singleRow => {
        // 仅处理 pxl 6 到 46 之间的行
        if (singleRow.pxl >= 6 && singleRow.pxl <= 46) {
          // 计算目标 pxl 并查找情况表中的对应行
          const targetPxl = singleRow.pxl + rowOffset;
          let distRow = distSheet.rows.find(row => row.pxl === targetPxl);

          if (!distRow) {
            distRow = { pxl: targetPxl };
            distSheet.rows.push(distRow);
          }

          if (singleRow.F3 !== undefined) {
            distRow[targetCol1] = singleRow.F3 || singleRow.F7;
          }

          if (singleRow.F4 !== undefined) {
            distRow[targetCol2] = singleRow.F4 || singleRow.F8;
          }

          const getNum = (val) => {
            if (val === undefined || val === null || val === "——" || val === "") return 0;
            const num = Number(val);
            return isNaN(num) ? 0 : num;
          };

          const f5 = getNum(distRow.F5);
          const f7 = getNum(distRow.F7);
          const f9 = getNum(distRow.F9);
          distRow.F3 = f5 + f7 + f9;

          const f6 = getNum(distRow.F6);
          const f8 = getNum(distRow.F8);
          const f10 = getNum(distRow.F10);
          distRow.F4 = f6 + f8 + f10;
        }
      });
    }
  }

  // 资产月报复核确认表逻辑
  if (confirmSheet) {
    // 定义单户表 pxl 到 确认表 pxl 的映射关系
    const confirmPxlMap = {
      6: 5,
      7: 6,
      8: 7,
      9: 8,
      10: 9,
      11: 10,
      12: 11,
      13: 12,
      16: 13,
      19: 14,
      20: 15,
      21: 16,
      25: 17,
      35: 18,
      36: 19,
      37: 20,
      41: 21,
      42: 22,
      43: 23
    };

    singleSheet.rows.forEach(singleRow => {
      const targetPxl = confirmPxlMap[singleRow.pxl];
      if (targetPxl !== undefined) {
        let confirmRow = confirmSheet.rows.find(row => row.pxl === targetPxl);

        if (!confirmRow) {
          confirmRow = { pxl: targetPxl };
          confirmSheet.rows.push(confirmRow);
        }

        // 将单户表的 F3, F4 填充到确认表的 F3, F4
        if (singleRow.F3 !== undefined) {
          confirmRow.F3 = singleRow.F3;
        }
        if (singleRow.F4 !== undefined) {
          confirmRow.F4 = singleRow.F4;
        }

        // 辅助函数：将值转换为数字，如果不能转换或未定义则返回 0，并区分完全为空的情况
        const getVal = (val) => {
          if (val === undefined || val === null || val === "——" || val === "") return { isEmpty: true, num: 0 };
          const num = Number(val);
          return isNaN(num) ? { isEmpty: true, num: 0 } : { isEmpty: false, num: num };
        };

        const c = getVal(confirmRow.F3);
        const d = getVal(confirmRow.F4);
        let result = "";

        if (c.num !== 0) {
          result = Math.round(((d.num - c.num) / Math.abs(c.num)) * 100 * 100) / 100;
        } else {
          if ((c.num === 0 && d.num === 0) || (c.isEmpty && d.isEmpty)) {
            result = 0;
          } else if (c.num === 0 && d.num !== 0) {
            result = Math.round((d.num / Math.abs(d.num)) * 100 * 100) / 100;
          }
        }

        if (result !== "") {
          // 若绝对值超过20则在后面加个*
          if (Math.abs(result) > 20) {
            confirmRow.F5 = `${result}*`;
          } else {
            confirmRow.F5 = result;
          }
        } else {
          confirmRow.F5 = "";
        }
      }
    });
  }

  return data;
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
        let finalData = JSON.parse(rawData);
        console.log("合并后数据对象类型:", typeof finalData);

        // 如果是保存纯数据，执行业务逻辑处理
        if (isPureData) {
          finalData = processPureData(finalData);
        }

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
