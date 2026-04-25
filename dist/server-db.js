const express = require("express");
const mysql = require("mysql2/promise");
const app = express();

app.use(express.json({ limit: "500mb" }));
app.use(express.urlencoded({ limit: "500mb", extended: true }));
app.use(express.text({ limit: "500mb" }));

app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Headers", "Content-Type");
  next();
});

const fs = require("fs");
const path = require("path");

// 确保数据目录存在
const DATA_DIR = path.join(__dirname, "data");
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

// 数据库配置
const dbConfig = {
  host: "piscesys.com",
  user: "spreadsheet",
  password: "Jiangge...0421",
  database: "spreadsheet",
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0
};

const pool = mysql.createPool(dbConfig);

// 辅助函数：通过报告ID获取各个表的 sheet_id
async function getSheetIds(connection, reportId) {
  const [rows] = await connection.query(
    "SELECT id, sheet_name FROM report_sheet WHERE report_id = ?",
    [reportId]
  );

  const sheetIds = {};
  rows.forEach(row => {
    if (row.sheet_name.includes("封面代码")) sheetIds.cover = row.id;
    if (row.sheet_name.includes("资产情况单户表")) sheetIds.single = row.id;
    if (row.sheet_name.includes("资产情况汇总表")) sheetIds.summary = row.id;
    if (row.sheet_name.includes("地方资产分布情况表")) sheetIds.dist = row.id;
    if (row.sheet_name.includes("资产月报复核确认表")) sheetIds.confirm = row.id;
  });

  return sheetIds;
}

// 辅助函数：从封面代码获取单位类型和层级
async function getCoverInfo(connection, coverSheetId) {
  const info = {
    isAdministrativeUnit: false,
    isPublicInstitution: false,
    isNonProfitOrganization: false,
    level: null
  };

  if (!coverSheetId) return info;

  const [rows] = await connection.query(
    "SELECT row_index, F4, F5, F8 FROM report_cell_data WHERE sheet_id = ? AND row_index IN (11, 13)",
    [coverSheetId]
  );

  rows.forEach(row => {
    if (row.row_index === 11) {
      info.isAdministrativeUnit = row.F4 === "1.行政单位";
      info.isPublicInstitution = row.F4 === "2.事业单位";
      info.isNonProfitOrganization = row.F4 === "3.民间非盈利组织";
    }
    if (row.row_index === 13) {
      info.level = row.F7;
    }
  });

  return info;
}

// 执行数据库层面的数据同步和计算逻辑
async function executeDatabaseLogic(connection, reportId) {
  const sheetIds = await getSheetIds(connection, reportId);
  const info = await getCoverInfo(connection, sheetIds.cover);

  // 1. 资产情况汇总表填充逻辑
  if (sheetIds.single && sheetIds.summary && (info.isAdministrativeUnit || info.isPublicInstitution || info.isNonProfitOrganization)) {
    let updateQuery = "";

    if (info.isAdministrativeUnit) {
      updateQuery = `
        UPDATE report_cell_data sum_tbl
        JOIN report_cell_data sin_tbl ON sum_tbl.row_index = sin_tbl.row_index 
        SET sum_tbl.F7 = sin_tbl.F3, sum_tbl.F8 = sin_tbl.F4
        WHERE sum_tbl.sheet_id = ? AND sin_tbl.sheet_id = ?
      `;
    } else if (info.isPublicInstitution) {
      updateQuery = `
        UPDATE report_cell_data sum_tbl
        JOIN report_cell_data sin_tbl ON sum_tbl.row_index = sin_tbl.row_index 
        SET sum_tbl.F11 = sin_tbl.F3, sum_tbl.F12 = sin_tbl.F4
        WHERE sum_tbl.sheet_id = ? AND sin_tbl.sheet_id = ?
      `;
    } else if (info.isNonProfitOrganization) {
      updateQuery = `
        UPDATE report_cell_data sum_tbl
        JOIN report_cell_data sin_tbl ON sum_tbl.row_index = sin_tbl.row_index 
        SET sum_tbl.F15 = sin_tbl.F7, sum_tbl.F16 = sin_tbl.F8
        WHERE sum_tbl.sheet_id = ? AND sin_tbl.sheet_id = ?
      `;
    }

    if (updateQuery) {
      await connection.query(updateQuery, [sheetIds.summary, sheetIds.single]);
    }

    // 计算汇总表每一行的 F3 (F7 + F11 + F15) 和 F4 (F8 + F12 + F16)
    await connection.query(`
      UPDATE report_cell_data
      SET F3 = IFNULL(CAST(NULLIF(F7, '——') AS DECIMAL(15,2)), 0) + 
               IFNULL(CAST(NULLIF(F11, '——') AS DECIMAL(15,2)), 0) + 
               IFNULL(CAST(NULLIF(F15, '——') AS DECIMAL(15,2)), 0),
          F4 = IFNULL(CAST(NULLIF(F8, '——') AS DECIMAL(15,2)), 0) + 
               IFNULL(CAST(NULLIF(F12, '——') AS DECIMAL(15,2)), 0) + 
               IFNULL(CAST(NULLIF(F16, '——') AS DECIMAL(15,2)), 0)
      WHERE sheet_id = ?
    `, [sheetIds.summary]);
  }

  // 2. 地方资产分布情况表逻辑
  if (sheetIds.single && sheetIds.dist && info.level && (info.isAdministrativeUnit || info.isPublicInstitution)) {
    let targetCol1, targetCol2;
    if (info.level === "2.省级") {
      targetCol1 = "F5"; targetCol2 = "F6";
    } else if (info.level === "3.地（市）级") {
      targetCol1 = "F7"; targetCol2 = "F8";
    } else if (info.level === "4.县级" || info.level === "5.乡镇级") {
      targetCol1 = "F9"; targetCol2 = "F10";
    }

    if (targetCol1 && targetCol2) {
      const rowOffset = info.isAdministrativeUnit ? 4 : (info.isPublicInstitution ? 48 : 0);

      await connection.query(`
        UPDATE report_cell_data dist_tbl
        JOIN report_cell_data sin_tbl ON dist_tbl.row_index = sin_tbl.row_index + ?
        SET dist_tbl.${targetCol1} = sin_tbl.F3, dist_tbl.${targetCol2} = sin_tbl.F4
        WHERE dist_tbl.sheet_id = ? AND sin_tbl.sheet_id = ? AND sin_tbl.row_index BETWEEN 6 AND 46
      `, [rowOffset, sheetIds.dist, sheetIds.single]);

      // 计算情况表每一行的 F3 (F5 + F7 + F9) 和 F4 (F6 + F8 + F10)
      await connection.query(`
        UPDATE report_cell_data
        SET F3 = IFNULL(CAST(NULLIF(F5, '——') AS DECIMAL(15,2)), 0) + 
                 IFNULL(CAST(NULLIF(F7, '——') AS DECIMAL(15,2)), 0) + 
                 IFNULL(CAST(NULLIF(F9, '——') AS DECIMAL(15,2)), 0),
            F4 = IFNULL(CAST(NULLIF(F6, '——') AS DECIMAL(15,2)), 0) + 
                 IFNULL(CAST(NULLIF(F8, '——') AS DECIMAL(15,2)), 0) + 
                 IFNULL(CAST(NULLIF(F10, '——') AS DECIMAL(15,2)), 0)
        WHERE sheet_id = ?
      `, [sheetIds.dist]);
    }
  }

  // 3. 资产月报复核确认表逻辑
  if (sheetIds.single && sheetIds.confirm) {
    // MySQL 中可以使用 CASE WHEN 进行复杂的映射更新
    await connection.query(`
      UPDATE report_cell_data conf_tbl
      JOIN report_cell_data sin_tbl ON 
        (sin_tbl.row_index = 6 AND conf_tbl.row_index = 5) OR
        (sin_tbl.row_index = 7 AND conf_tbl.row_index = 6) OR
        (sin_tbl.row_index = 8 AND conf_tbl.row_index = 7) OR
        (sin_tbl.row_index = 9 AND conf_tbl.row_index = 8) OR
        (sin_tbl.row_index = 10 AND conf_tbl.row_index = 9) OR
        (sin_tbl.row_index = 11 AND conf_tbl.row_index = 10) OR
        (sin_tbl.row_index = 12 AND conf_tbl.row_index = 11) OR
        (sin_tbl.row_index = 13 AND conf_tbl.row_index = 12) OR
        (sin_tbl.row_index = 16 AND conf_tbl.row_index = 13) OR
        (sin_tbl.row_index = 19 AND conf_tbl.row_index = 14) OR
        (sin_tbl.row_index = 20 AND conf_tbl.row_index = 15) OR
        (sin_tbl.row_index = 21 AND conf_tbl.row_index = 16) OR
        (sin_tbl.row_index = 25 AND conf_tbl.row_index = 17) OR
        (sin_tbl.row_index = 35 AND conf_tbl.row_index = 18) OR
        (sin_tbl.row_index = 36 AND conf_tbl.row_index = 19) OR
        (sin_tbl.row_index = 37 AND conf_tbl.row_index = 20) OR
        (sin_tbl.row_index = 40 AND conf_tbl.row_index = 21) OR
        (sin_tbl.row_index = 42 AND conf_tbl.row_index = 22) OR
        (sin_tbl.row_index = 43 AND conf_tbl.row_index = 23)
      SET conf_tbl.F3 = sin_tbl.F3, conf_tbl.F4 = sin_tbl.F4
      WHERE conf_tbl.sheet_id = ? AND sin_tbl.sheet_id = ?
    `, [sheetIds.confirm, sheetIds.single]);

    // 环比增长计算
    // 在 SQL 中进行复杂的计算
    await connection.query(`
      UPDATE report_cell_data target
      JOIN (
        SELECT 
          id,
          CASE 
            WHEN val_c != 0 THEN 
              CASE 
                WHEN ABS(ROUND(((val_d - val_c) / ABS(val_c)) * 100, 2)) > 20 
                THEN CONCAT(ROUND(((val_d - val_c) / ABS(val_c)) * 100, 2), '*')
                ELSE CAST(ROUND(((val_d - val_c) / ABS(val_c)) * 100, 2) AS CHAR)
              END
            WHEN val_c = 0 AND val_d = 0 THEN '0'
            WHEN val_c = 0 AND val_d != 0 THEN 
              CASE 
                WHEN ABS(ROUND((val_d / ABS(val_d)) * 100, 2)) > 20 
                THEN CONCAT(ROUND((val_d / ABS(val_d)) * 100, 2), '*')
                ELSE CAST(ROUND((val_d / ABS(val_d)) * 100, 2) AS CHAR)
              END
            ELSE ''
          END as calc_result
        FROM (
          SELECT 
            id,
            IFNULL(CAST(NULLIF(TRIM(F3), '——') AS DECIMAL(15,2)), 0) AS val_c,
            IFNULL(CAST(NULLIF(TRIM(F4), '——') AS DECIMAL(15,2)), 0) AS val_d
          FROM report_cell_data
          WHERE sheet_id = ? AND row_index IN (5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23)
        ) AS temp
      ) AS calc_data ON target.id = calc_data.id
      SET target.F5 = calc_data.calc_result
    `, [sheetIds.confirm]);
  }
}

// 在内存中暂存分片数据（实际生产中应使用 Redis 或其他临时存储）
const chunksStorage = {};

// 保存分片数据 (改为保存到数据库)
app.post("/api/chunk-upload", async (req, res) => {
  const { chunkIndex, totalChunks, id, pure, do: action, template_id } = req.query;
  const data = req.body;
  const tempDir = path.join(__dirname, "temp");

  // 确保临时目录存在
  if (!fs.existsSync(tempDir)) {
    fs.mkdirSync(tempDir);
  }

  const chunkPath = path.join(tempDir, `${id}-${chunkIndex}`);
  fs.writeFileSync(chunkPath, JSON.stringify(data));

  if (!chunksStorage[id]) {
    chunksStorage[id] = new Array(parseInt(totalChunks, 10));
  }
  chunksStorage[id][parseInt(chunkIndex, 10)] = data;

  if (pure === "true") {
    console.log("接收到最后一个分片，开始处理数据入库...");

    try {
      // 合并分片数据
      let rawData = "";
      for (let i = 0; i < chunksStorage[id].length; i++) {
        if (chunksStorage[id][i] && chunksStorage[id][i].data) {
          rawData += chunksStorage[id][i].data;
        }
      }

      const finalData = JSON.parse(rawData);
      delete chunksStorage[id]; // 清理内存

      // 获取数据库连接并开启事务
      const connection = await pool.getConnection();
      await connection.beginTransaction();

      try {
        let newTemplateId = null;
        if (action === "SaveReportTemplate") {
          // 1. 如果 template_id 还没有包含 report_id，加上 report_id 作为专属模板名
          newTemplateId = template_id;
          if (!newTemplateId.endsWith(`_${id}`)) {
            newTemplateId = `${template_id}_${id}`;
          }

          // 2. 存入本地文件
          const finalPath = getFileName(newTemplateId, false);
          fs.writeFileSync(finalPath, JSON.stringify(finalData, null, 2));
          const relativePath = path.relative(__dirname, finalPath).replace(/\\/g, '/');

          // 3. 将模板路径存入数据库
          await connection.query(
            "INSERT INTO report_template (id, template_path) VALUES (?, ?) ON DUPLICATE KEY UPDATE template_path = VALUES(template_path)",
            [newTemplateId, relativePath]
          );

          // 4. 更新 report_main 关联
          await connection.query(
            "UPDATE report_main SET template_id = ? WHERE id = ?",
            [newTemplateId, id]
          );
        } else if (action === "SavePureData") {
          // 如果是纯数据（保存报表数据）
          // 1. 插入或更新 report_main
          // 如果前端传了 template_id，则更新关联关系
          if (template_id) {
            await connection.query(
              "INSERT INTO report_main (id, template_id) VALUES (?, ?) ON DUPLICATE KEY UPDATE template_id = VALUES(template_id), updated_at = CURRENT_TIMESTAMP",
              [id, template_id]
            );
          } else {
            await connection.query(
              "INSERT INTO report_main (id) VALUES (?) ON DUPLICATE KEY UPDATE updated_at = CURRENT_TIMESTAMP",
              [id]
            );
          }

          // 获取 dwdm
          let dwdm = null;
          const [mainRows] = await connection.query("SELECT dwdm FROM report_main WHERE id = ?", [id]);
          if (mainRows.length > 0) {
            dwdm = mainRows[0].dwdm;
          }

          const reportInfoData = {};

          // 2. 遍历 sheets
          for (const sheet of finalData) {
            if (!sheet.name) continue;

            // 提取有ID的表和单元格数据
            if (sheet.sheetId) {
              reportInfoData['SheetID'] = sheet.sheetId; // 字段名为 SheetID，内容为设置的表ID
            }

            if (sheet.rows && Array.isArray(sheet.rows)) {
              for (const row of sheet.rows) {
                for (let i = 1; i <= 20; i++) {
                  const idKey = `id_${i}`;
                  const valKey = `F${i}`;
                  if (row[idKey] && row[valKey] !== undefined) {
                    reportInfoData[row[idKey]] = row[valKey];
                  }
                }
              }
            }

            // 然后有ID这个表就不保存到celldata表
            if (sheet.sheetId) {
              continue;
            }

            // 查询 sheet 是否已存在
            const [sheetRows] = await connection.query(
              "SELECT id FROM report_sheet WHERE report_id = ? AND sheet_name = ?",
              [id, sheet.name]
            );

            let sheetId;
            if (sheetRows.length > 0) {
              sheetId = sheetRows[0].id;
            } else {
              // 不存在则插入
              const [result] = await connection.query(
                "INSERT INTO report_sheet (report_id, sheet_name) VALUES (?, ?)",
                [id, sheet.name]
              );
              sheetId = result.insertId;
            }

            // 3. 处理单元格数据
            if (sheet.rows && Array.isArray(sheet.rows)) {
              for (const row of sheet.rows) {
                if (row.pxl === undefined) continue;

                // 构建 UPDATE 和 INSERT 语句的数据
                const fields = ["sheet_id", "row_index"];
                const values = [sheetId, row.pxl];
                const updatePairs = [];

                for (let i = 1; i <= 20; i++) {
                  const key = `F${i}`;
                  if (row[key] !== undefined) {
                    fields.push(key);
                    values.push(row[key]);
                    updatePairs.push(`${key} = VALUES(${key})`);
                  }
                }

                if (fields.length > 2) {
                  const placeholders = new Array(fields.length).fill("?").join(", ");
                  let query = `INSERT INTO report_cell_data (${fields.join(", ")}) VALUES (${placeholders}) ON DUPLICATE KEY UPDATE ${updatePairs.join(", ")}`;

                  await connection.query(query, values);
                }
              }
            }
          }

          // 保存提取出的 info 数据到 report_info 表
          if (dwdm && Object.keys(reportInfoData).length > 0) {
            // 分离 bsrq 字段，将其更新到 report_main，并从 reportInfoData 中移除
            if (reportInfoData['bsrq'] !== undefined) {
              const bsrqVal = reportInfoData['bsrq'];
              delete reportInfoData['bsrq'];

              // 更新 main 表的 reported_at
              try {
                await connection.query("UPDATE report_main SET reported_at = ? WHERE id = ?", [bsrqVal, id]);
              } catch (e) {
                console.error("更新 report_main reported_at 失败:", e.message);
              }
            }

            // 如果移除 bsrq 后还有其他字段，则继续更新 report_info
            if (Object.keys(reportInfoData).length > 0) {
              const keys = Object.keys(reportInfoData);
              const values = Object.values(reportInfoData);

              let currentSheetId = reportInfoData['SheetID'] || "";
              if (!reportInfoData.hasOwnProperty('SheetID')) {
                keys.push('SheetID');
                values.push(currentSheetId);
              }

              const fieldsStr = ["dwdm", ...keys].map(k => `\`${k}\``).join(", ");
              const placeholders = new Array(keys.length + 1).fill("?").join(", ");
              const updatePairs = keys.map(k => `\`${k}\` = VALUES(\`${k}\`)`);
              const vals = [dwdm, ...values];

              const infoQuery = `INSERT INTO report_info (${fieldsStr}) VALUES (${placeholders}) ON DUPLICATE KEY UPDATE ${updatePairs.join(", ")}`;

              try {
                await connection.query(infoQuery, vals);
              } catch (e) {
                console.error("保存 report_info 失败 (请确保数据库表及字段已创建，并且存在 dwdm 和 SheetID 联合唯一索引):", e.message);
              }
            }
          }
        } else {
          // 默认 SaveXspreadSheet（即保存模板）
          // 1. 存入本地文件
          const finalPath = getFileName(template_id, false);
          fs.writeFileSync(finalPath, JSON.stringify(finalData, null, 2));
          const relativePath = path.relative(__dirname, finalPath).replace(/\\/g, '/');

          // 2. 将模板路径存入数据库
          // 使用传入的 template_id，如果没有则使用区分 id
          const tid = template_id || id;
          await connection.query(
            "INSERT INTO report_template (id, template_path) VALUES (?, ?) ON DUPLICATE KEY UPDATE template_path = VALUES(template_path)",
            [tid, relativePath]
          );
        }

        if (template_id === "yuebaoIndex3432") {
          console.log("执行数据库自动填充和计算逻辑...");
          await executeDatabaseLogic(connection, id);
        }

        await connection.commit();
        console.log("数据处理完成");

        res.json({
          success: true,
          chunkIndex,
          merged: true,
          new_template_id: newTemplateId,
          message: `Chunk ${chunkIndex} of ${totalChunks} received and data processed successfully`
        });
      } catch (error) {
        await connection.rollback();
        throw error;
      } finally {
        connection.release();
      }
    } catch (error) {
      console.error("处理错误:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  } else {
    res.json({
      success: true,
      chunkIndex,
      message: `Chunk ${chunkIndex} of ${totalChunks} received`,
    });
  }
});

// 获取表格总数 (从模板文件获取)
app.get("/api/getSheetCount", async (req, res) => {
  const { id, template_id } = req.query;

  try {
    let filePath;

    if (template_id) {
      // 优先从传入的 template_id 查找
      const [rows] = await pool.query(
        "SELECT template_path FROM report_template WHERE id = ?",
        [template_id]
      );
      if (rows.length > 0 && rows[0].template_path) {
        filePath = path.join(__dirname, rows[0].template_path);
      }
    }

    if (!filePath) {
      // 从主表关联的 template_id 查找
      const [rows] = await pool.query(
        "SELECT t.template_path FROM report_main m JOIN report_template t ON m.template_id = t.id WHERE m.id = ?",
        [id]
      );
      if (rows.length > 0 && rows[0].template_path) {
        filePath = path.join(__dirname, rows[0].template_path);
      }
    }

    if (!filePath) {
      // 兼容旧逻辑
      filePath = getFileName(template_id || id, false);
    }

    if (fs.existsSync(filePath)) {
      const data = JSON.parse(fs.readFileSync(filePath, "utf8"));
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

// 获取单个表格数据 (从模板文件获取)
app.get("/api/getSheet", async (req, res) => {
  const { id, template_id, index } = req.query;

  try {
    let filePath;

    if (template_id) {
      // 优先从传入的 template_id 查找
      const [rows] = await pool.query(
        "SELECT template_path FROM report_template WHERE id = ?",
        [template_id]
      );
      if (rows.length > 0 && rows[0].template_path) {
        filePath = path.join(__dirname, rows[0].template_path);
      }
    }

    if (!filePath) {
      // 从主表关联的 template_id 查找
      const [rows] = await pool.query(
        "SELECT t.template_path FROM report_main m JOIN report_template t ON m.template_id = t.id WHERE m.id = ?",
        [id]
      );
      if (rows.length > 0 && rows[0].template_path) {
        filePath = path.join(__dirname, rows[0].template_path);
      }
    }

    if (!filePath) {
      // 兼容旧逻辑
      filePath = getFileName(template_id || id, false);
    }

    if (fs.existsSync(filePath)) {
      const data = JSON.parse(fs.readFileSync(filePath, "utf8"));
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
// 从数据库获取纯数据
app.get("/api/getPureData", async (req, res) => {
  const { id } = req.query;

  try {
    // 1. 获取 dwdm, template_id, reported_at
    let dwdm = null;
    let template_id = null;
    let reported_at = null;
    const [mainRows] = await pool.query("SELECT dwdm, template_id, reported_at FROM report_main WHERE id = ?", [id]);
    if (mainRows.length > 0) {
      dwdm = mainRows[0].dwdm;
      template_id = mainRows[0].template_id;
      reported_at = mainRows[0].reported_at;
    }

    // 2. 获取 report_info 数据
    let infoDataList = [];
    if (dwdm) {
      try {
        const [infoRows] = await pool.query("SELECT * FROM report_info WHERE dwdm = ?", [dwdm]);
        infoDataList = infoRows; // 可能包含多条记录（不同的 SheetID）

        // 注入 bsrq 字段
        if (reported_at) {
          infoDataList.forEach(info => {
            // 假设需要格式化日期为 YYYY-MM-DD
            const d = new Date(reported_at);
            info.bsrq = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
          });
        }
      } catch (e) {
        console.error("读取 report_info 失败:", e.message);
      }
    }

    // 3. 读取模板数据，用于匹配 sheetId 和单元格 id
    let templateData = [];
    try {
      let filePath;
      if (template_id) {
        const [rows] = await pool.query("SELECT template_path FROM report_template WHERE id = ?", [template_id]);
        if (rows.length > 0 && rows[0].template_path) {
          filePath = path.join(__dirname, rows[0].template_path);
        }
      }
      if (!filePath) {
        filePath = getFileName(template_id || id, false);
      }
      if (fs.existsSync(filePath)) {
        templateData = JSON.parse(fs.readFileSync(filePath, "utf8"));
      }
    } catch (e) {
      console.error("读取模板失败:", e.message);
    }

    const data = [];

    // 4. 根据模板结构，结合 infoDataList 和 report_cell_data 生成纯数据
    for (const tSheet of templateData) {
      const sheetData = {
        name: tSheet.name,
        rows: []
      };

      // 从 infoDataList 找出匹配当前模板 sheetId 的那条记录
      let matchedInfo = null;
      if (tSheet.sheetId) {
        matchedInfo = infoDataList.find(r => r.SheetID === tSheet.sheetId);
      } else {
        matchedInfo = infoDataList.find(r => !r.SheetID) || infoDataList[0] || {};
      }
      const infoData = matchedInfo || {};

      let hasInfoData = false;

      // 如果模板表有 sheetId 且和 infoData 中的 SheetID 匹配，则说明这个表之前存过 ID
      if (tSheet.sheetId && infoData.SheetID === tSheet.sheetId) {
        hasInfoData = true;
      }
      if (!hasInfoData) {
        const [sheets] = await pool.query(
          "SELECT id FROM report_sheet WHERE report_id = ? AND sheet_name = ?",
          [id, tSheet.name]
        );

        if (sheets.length > 0) {
          const [rows] = await pool.query(
            "SELECT * FROM report_cell_data WHERE sheet_id = ? ORDER BY row_index ASC",
            [sheets[0].id]
          );

          rows.forEach(row => {
            const rowData = { pxl: row.row_index };
            for (let i = 1; i <= 20; i++) {
              const key = `F${i}`;
              if (row[key] !== null) {
                rowData[key] = row[key];
              }
            }
            sheetData.rows.push(rowData);
          });
        }
      }

      // 处理单元格级别的 ID 填充
      if (tSheet.rows) {
        Object.entries(tSheet.rows).forEach(([rowIndexStr, row]) => {
          if (row && row.cells) {
            Object.entries(row.cells).forEach(([colIndexStr, cell]) => {
              if (cell && cell.id && infoData[cell.id] !== undefined) {
                // 确保对应的行存在
                const pxl = parseInt(rowIndexStr) + 1;
                const colIndex = parseInt(colIndexStr) + 1;
                let targetRow = sheetData.rows.find(r => r.pxl === pxl);
                if (!targetRow) {
                  targetRow = { pxl: pxl };
                  sheetData.rows.push(targetRow);
                }
                targetRow[`F${colIndex}`] = infoData[cell.id];
              }
            });
          }
        });
      }

      if (sheetData.rows.length > 0) {
        // 对行排序以防乱序
        sheetData.rows.sort((a, b) => a.pxl - b.pxl);
        data.push(sheetData);
      }
    }

    res.json({
      success: true,
      data: data,
      infoDataList: infoDataList
    });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 获取报表列表
app.get("/api/getReportList", async (req, res) => {
  const { dwdm } = req.query;

  if (!dwdm) {
    return res.json({ success: false, message: "缺少单位代码 dwdm 参数" });
  }

  try {
    const [rows] = await pool.query(
      "SELECT id, template_id, report_name, report_year, report_month, isReported, created_at, updated_at FROM report_main WHERE dwdm = ? ORDER BY updated_at DESC",
      [dwdm]
    );

    res.json({
      success: true,
      data: rows
    });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 新增报表
app.post("/api/createReport", async (req, res) => {
  const { id, dwdm, report_name, report_year, report_month, template_id, inherit_last_month } = req.body;

  if (!id || !dwdm || !report_name || !report_year || !report_month || !template_id) {
    return res.json({ success: false, message: "缺少必要参数" });
  }

  const connection = await pool.getConnection();
  await connection.beginTransaction();

  try {
    // 1. 在 report_main 表中创建新记录
    await connection.query(
      "INSERT INTO report_main (id, template_id, report_name, report_year, report_month, dwdm) VALUES (?, ?, ?, ?, ?, ?)",
      [id, template_id, report_name, report_year, report_month, dwdm]
    );

    // 2. 如果开启了继承上月数据，执行继承逻辑
    if (inherit_last_month) {
      // 计算上个月
      let lastMonth = report_month - 1;
      let lastYear = report_year;
      if (lastMonth === 0) {
        lastMonth = 12;
        lastYear = report_year - 1;
      }

      // 查找上月同一单位、同一模板的最新报表记录
      const [lastReportRows] = await connection.query(
        "SELECT id FROM report_main WHERE dwdm = ? AND template_id = ? AND report_year = ? AND report_month = ? ORDER BY updated_at DESC LIMIT 1",
        [dwdm, template_id, lastYear, lastMonth]
      );

      if (lastReportRows.length > 0) {
        const lastReportId = lastReportRows[0].id;

        // 获取上个月报表的 sheets
        const [lastSheets] = await connection.query(
          "SELECT id, sheet_name FROM report_sheet WHERE report_id = ?",
          [lastReportId]
        );

        for (const lastSheet of lastSheets) {
          // 在当前报表下创建相同名字的 sheet
          const [insertSheetResult] = await connection.query(
            "INSERT INTO report_sheet (report_id, sheet_name) VALUES (?, ?)",
            [id, lastSheet.sheet_name]
          );
          const newSheetId = insertSheetResult.insertId;

          if (lastSheet.sheet_name.includes("封面代码")) {
            // 封面代码完全复制
            await connection.query(`
              INSERT INTO report_cell_data 
                (sheet_id, row_index, F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12, F13, F14, F15, F16, F17, F18, F19, F20)
              SELECT 
                ?, row_index, F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12, F13, F14, F15, F16, F17, F18, F19, F20
              FROM report_cell_data 
              WHERE sheet_id = ?
            `, [newSheetId, lastSheet.id]);
          } else if (lastSheet.sheet_name.includes("资产情况单户表")) {
            // 单户表：将上月的 F4 复制到本月的 F3
            await connection.query(`
              INSERT INTO report_cell_data 
                (sheet_id, row_index, F3)
              SELECT 
                ?, row_index, F4
              FROM report_cell_data 
              WHERE sheet_id = ? AND F4 IS NOT NULL
            `, [newSheetId, lastSheet.id]);
          }
        }
        console.log(`成功从上月报表 ${lastReportId} 继承了数据到本月报表 ${id}`);
      } else {
        console.log(`未找到上个月(${lastYear}-${lastMonth})的数据，跳过继承`);
      }
    }

    await connection.commit();
    res.json({ success: true, message: "报表创建成功", id: id });
  } catch (error) {
    await connection.rollback();
    console.error("创建报表失败:", error);
    res.status(500).json({ success: false, message: "创建报表失败", error: error.message });
  } finally {
    connection.release();
  }
});

// 校验报表接口
app.get("/api/verifyReport", async (req, res) => {
  const { id } = req.query;

  if (!id) {
    return res.json({ success: false, message: "缺少报表ID参数" });
  }

  const errors = [];
  try {
    const sheetIds = await getSheetIds(pool, id);

    if (sheetIds.single) {
      // 1. 验证单户表 pxl6-46，F4列是否有填写数据
      const [rows1] = await pool.query(`
        SELECT row_index FROM report_cell_data 
        WHERE sheet_id = ? AND row_index BETWEEN 6 AND 46 
        AND (F4 IS NULL OR TRIM(F4) = '')
      `, [sheetIds.single]);

      if (rows1.length > 0) {
        const rowIndexes = rows1.map(r => r.row_index).join(", ");
        errors.push(`资产情况单户表：第 ${rowIndexes} 行未填写数据，请填报本月期末数`);
      }

      // 2. 验证单户表 pxl44+45+46 的 F4 是否等于 pxl43 的 F4
      const [rows2] = await pool.query(`
        SELECT row_index, IFNULL(CAST(NULLIF(TRIM(F4), '——') AS DECIMAL(15,2)), 0) as val 
        FROM report_cell_data 
        WHERE sheet_id = ? AND row_index IN (43, 44, 45, 46)
      `, [sheetIds.single]);

      const valMap = { 43: 0, 44: 0, 45: 0, 46: 0 };
      rows2.forEach(r => valMap[r.row_index] = Number(r.val));

      // 使用 Math.round 避免浮点数精度问题，或者直接用 Number 相加
      const sum = Number((valMap[44] + valMap[45] + valMap[46]).toFixed(2));
      const targetVal = Number(valMap[43].toFixed(2));

      if (sum !== targetVal) {
        errors.push(`资产情况单户表：累计盈余(${valMap[44]}) + 专用基金(${valMap[45]}) + 本期盈余(${valMap[46]}) = ${sum}，必须等于 净资产合计(${targetVal})`);
      }
    }

    if (sheetIds.confirm) {
      // 3. 确认表环比增长 F5列中绝对值大于20则F6列必填
      const [rows3] = await pool.query(`
        SELECT row_index, F5, F6 
        FROM report_cell_data 
        WHERE sheet_id = ? AND F5 LIKE '%*%' AND (F6 IS NULL OR TRIM(F6) = '')
      `, [sheetIds.confirm]);

      if (rows3.length > 0) {
        rows3.forEach(r => {
          errors.push(`资产月报复核确认表：第 ${r.row_index} 行本期期末数相比上月期末数环比增长绝对值超过20%，必须填写备注`);
        });
      }
    }

    if (errors.length > 0) {
      res.json({ success: true, valid: false, errors: errors });
    } else {
      res.json({ success: true, valid: true });
    }
  } catch (error) {
    console.error("校验报表失败:", error);
    res.status(500).json({ success: false, message: "服务端校验执行异常", error: error.message });
  }
});

// 提交报表接口
app.get("/api/submitReport", async (req, res) => {
  const { id } = req.query;

  if (!id) {
    return res.json({ success: false, message: "缺少报表ID参数" });
  }

  try {
    const [rows] = await pool.query("SELECT reported_at FROM report_main WHERE id = ?", [id]);
    let reportedAtUpdate = "";
    let params = [id];

    if (rows.length > 0) {
      if (!rows[0].reported_at) {
        // 如果为空，设置当前日期为报送时间
        const d = new Date();
        const dateStr = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
        reportedAtUpdate = ", reported_at = ?";
        params = [dateStr, id];
      }
    }

    const query = `UPDATE report_main SET isReported = true${reportedAtUpdate} WHERE id = ?`;
    await pool.query(query, params);

    res.json({ success: true, message: "上报成功" });
  } catch (error) {
    console.error("上报报表失败:", error);
    res.status(500).json({ success: false, message: "服务端上报执行异常", error: error.message });
  }
});

// 撤回报表接口
app.get("/api/withdrawReport", async (req, res) => {
  const { id } = req.query;

  if (!id) {
    return res.json({ success: false, message: "缺少报表ID参数" });
  }

  try {
    await pool.query("UPDATE report_main SET isReported = false WHERE id = ?", [id]);
    res.json({ success: true, message: "撤回成功" });
  } catch (error) {
    console.error("撤回报表失败:", error);
    res.status(500).json({ success: false, message: "服务端撤回执行异常", error: error.message });
  }
});

app.listen(3000, () => {
  console.log("Server running on port 3000 with Database configuration");
});