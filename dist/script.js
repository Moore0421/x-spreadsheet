// 请求网址
const url = "localhost";

// 显示加载动画
function showLoading() {
  document.getElementById("loadingOverlay").style.display = "flex";
}

// 隐藏加载动画
function hideLoading() {
  document.getElementById("loadingOverlay").style.display = "none";
}

// 安全移除progressDiv
function removeProgressDiv() {
  if (progressDiv && progressDiv.parentNode === document.body) {
    document.body.removeChild(progressDiv);
  }
}

// 从sessionStorage或URL获取参数
async function getUrlParam(name) {
  const sessionVal = sessionStorage.getItem(name);
  if (sessionVal !== null && sessionVal !== undefined) {
    return sessionVal;
  }
  var urlParams = new URLSearchParams(window.location.search);
  return urlParams.get(name);
}

// 格式化日期
function formatDate(dateStr) {
  if (!dateStr) return "-";
  const d = new Date(dateStr);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
}

// 加载报表列表
async function loadReportList() {
  const urlParams = new URLSearchParams(window.location.search);
  const mode = urlParams.get("mode");
  const templateId = urlParams.get("template_id");
  console.log(mode);

  if (mode) {
    sessionStorage.setItem("mode", mode);
    document.getElementById("report-list-container").style.display = "none";
    document.getElementById("x-spreadsheet-demo").style.display = "block";
    document.getElementById("app-top-bar-design").style.display = "block";
    document.getElementById("design-title").innerHTML = `设计报表 - ${templateId}`;
    await load();
    return;
  }

  const dwdm = await getUrlParam("dwdm");
  if (!dwdm) {
    alert("缺少单位代码(dwdm)参数");
    return;
  }

  sessionStorage.setItem("dwdm", dwdm);
  showLoading();

  // 检查是否是管理员模式，控制新增按钮显示
  const isAdmin = await getUrlParam("admin");
  if (isAdmin === "true") {
    document.getElementById("create-report-btn").style.display = "block";
    initCreateReportModal(dwdm);
  }

  try {
    const response = await fetch(`http://${url}:3000/api/getReportList?dwdm=${dwdm}`);
    const result = await response.json();

    if (result.success) {
      const tbody = document.getElementById("report-table-body");
      tbody.innerHTML = "";

      if (result.data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;color:#909399;">暂无报表数据</td></tr>`;
      } else {
        result.data.forEach(report => {
          const tr = document.createElement("tr");

          tr.innerHTML = `
            <td>${report.report_name || '-'}</td>
            <td>${report.report_year || '-'}</td>
            <td>${report.report_month || '-'}</td>
            <td>${formatDate(report.created_at)}</td>
            <td>${formatDate(report.updated_at)}</td>
            <td>
              <button class="action-btn edit-btn" data-id="${report.id}" data-template="${report.template_id || ''}" data-name="${report.report_name || '未命名报表'}" data-reported="${report.isReported ? 'true' : 'false'}">前往编报</button>
            </td>
          `;

          tbody.appendChild(tr);
        });

        // 绑定按钮事件
        document.querySelectorAll(".edit-btn").forEach(btn => {
          btn.addEventListener("click", function () {
            const id = this.getAttribute("data-id");
            const templateId = this.getAttribute("data-template");
            const reportName = this.getAttribute("data-name");
            const isReported = this.getAttribute("data-reported") === 'true';

            const enterMode = isReported ? "preview" : "enabled";

            // 隐式传参并切换界面
            enterReportEditor(id, templateId, enterMode, reportName, isReported);
          });
        });
      }
    } else {
      alert("获取报表列表失败: " + result.message);
    }
  } catch (error) {
    console.error("加载报表列表失败:", error);
    alert("加载报表列表失败");
  } finally {
    hideLoading();
  }
}

// 初始化新增报表弹窗事件
function initCreateReportModal(dwdm) {
  const modal = document.getElementById("createReportModal");
  const btn = document.getElementById("create-report-btn");
  const closeBtn = document.getElementById("closeModalBtn");
  const cancelBtn = document.getElementById("cancelModalBtn");
  const confirmBtn = document.getElementById("confirmCreateBtn");

  const now = new Date();
  document.getElementById("reportYear").value = now.getFullYear();
  document.getElementById("reportMonth").value = now.getMonth() + 1;

  const hideModal = () => {
    modal.style.display = "none";
    document.getElementById("reportName").value = "";
    document.getElementById("templateName").value = "";
    document.getElementById("inheritLastMonth").checked = true;
  };

  btn.onclick = (event) => {
    // 阻止事件冒泡，防止触发其他绑定的外层事件
    event.stopPropagation();
    modal.style.display = "flex";
  };

  closeBtn.onclick = hideModal;
  cancelBtn.onclick = hideModal;

  window.onclick = (event) => {
    if (event.target === modal) {
      hideModal();
    }
  };

  confirmBtn.onclick = async () => {
    const reportName = document.getElementById("reportName").value.trim();
    const reportYear = document.getElementById("reportYear").value;
    const reportMonth = document.getElementById("reportMonth").value;
    const templateId = document.getElementById("templateName").value.trim();
    const inheritLastMonth = document.getElementById("inheritLastMonth").checked;

    if (!reportName || !reportYear || !reportMonth || !templateId) {
      alert("请填写完整的报表信息！");
      return;
    }

    // 随机生成ID，例如 report_1234567890
    const newId = "report_" + Date.now() + Math.floor(Math.random() * 1000);

    const payload = {
      id: newId,
      dwdm: dwdm,
      report_name: reportName,
      report_year: parseInt(reportYear),
      report_month: parseInt(reportMonth),
      template_id: templateId,
      inherit_last_month: inheritLastMonth
    };

    try {
      confirmBtn.disabled = true;
      confirmBtn.textContent = "创建中...";

      const response = await fetch(`http://${url}:3000/api/createReport`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      const result = await response.json();

      if (result.success) {
        hideModal();
        alert("报表创建成功！");
        // 重新加载列表
        await loadReportList();
      } else {
        alert("报表创建失败: " + result.message);
      }
    } catch (error) {
      console.error("创建请求失败:", error);
      alert("创建报表请求失败！");
    } finally {
      confirmBtn.disabled = false;
      confirmBtn.textContent = "确认新增";
    }
  };
}

// 进入报表编辑器界面
async function enterReportEditor(id, templateId, mode, reportName, isReported = false) {
  // 存入 sessionStorage
  sessionStorage.setItem("id", id);
  sessionStorage.setItem("template_id", templateId);
  sessionStorage.setItem("mode", mode);
  sessionStorage.setItem("report_name", reportName || "未命名报表");
  sessionStorage.setItem("isReported", isReported ? "true" : "false");
  sessionStorage.setItem("verify_status", ""); // 每次进入清空校验状态

  // 切换UI显示
  document.getElementById("report-list-container").style.display = "none";
  document.getElementById("app-top-bar").style.display = "flex";
  document.getElementById("x-spreadsheet-demo").style.display = "block";

  // 动态设置标题
  document.getElementById("dynamic-report-title").textContent = reportName || "未命名报表";
  document.getElementById("x-spreadsheet-demo").innerHTML = "";
  await load();
}

// 上传分片
async function uploadChunk(
  chunk,
  chunkIndex,
  totalChunks,
  id,
  isLastChunk,
  action = "SaveXspreadSheet",
  template_id = null
) {
  // 包装数据为对象格式
  const payload = {
    data: chunk,
  };

  let requestUrl = `http://${url}:3000/api/chunk-upload?id=${id}&chunkIndex=${chunkIndex}&totalChunks=${totalChunks}&pure=${isLastChunk}&do=${action}`;

  if (template_id) {
    requestUrl += `&template_id=${template_id}`;
  }

  const response = await fetch(
    requestUrl,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload), // 发送包装后的数据
    }
  );
  return await response.json();
}

// 通用提示
const progressDiv = document.createElement("div");
progressDiv.style.position = "fixed";
progressDiv.style.top = "50%";
progressDiv.style.left = "50%";
progressDiv.style.transform = "translate(-50%, -50%)";
progressDiv.style.padding = "20px";
progressDiv.style.background = "rgba(0,0,0,0.7)";
progressDiv.style.color = "white";
progressDiv.style.borderRadius = "5px";

// 保存数据到服务器
async function saveToServer(data) {
  try {
    const id = (await getUrlParam("id")) || 0;
    const template_id = await getUrlParam("template_id");
    let chunkSize = 256 * 256;
    const jsonData = JSON.stringify(data);
    const totalChunks = Math.ceil(jsonData.length / chunkSize);
    // 显示进度提示
    document.body.appendChild(progressDiv);
    // 上传所有分片
    for (let i = 0; i < totalChunks; i++) {
      const chunk = jsonData.slice(i * chunkSize, (i + 1) * chunkSize);
      progressDiv.textContent = "正在上传 " + (i + 1) + " / " + totalChunks;
      // 判断是否是最后一个分片
      const isLastChunk = i === totalChunks - 1;
      const result = await uploadChunk(
        chunk,
        i,
        totalChunks,
        id,
        isLastChunk,
        "SaveXspreadSheet",
        template_id
      );
      if (!result.success) {
        progressDiv.textContent = "分片上传失败";
      }
      // 如果是最后一个分片且合并成功
      if (isLastChunk && result.merged) {
        progressDiv.textContent = "保存成功";
        await new Promise((resolve) => setTimeout(resolve, 2000));
        removeProgressDiv();
      }
    }
  } catch (error) {
    console.error("保存失败:", error);
    progressDiv.textContent = "保存失败: " + error.message;
    await new Promise((resolve) => setTimeout(resolve, 2000));
    removeProgressDiv();
  }
}

// 保存纯数据到服务器
async function savePureDataToServer(data) {
  try {
    const id = (await getUrlParam("id")) || 0;
    const template_id = await getUrlParam("template_id");
    const chunkSize = 256 * 256;
    const jsonData = JSON.stringify(data);
    const totalChunks = Math.ceil(jsonData.length / chunkSize);
    // 显示进度提示
    document.body.appendChild(progressDiv);
    // 上传所有分片
    for (let i = 0; i < totalChunks; i++) {
      const chunk = jsonData.slice(i * chunkSize, (i + 1) * chunkSize);
      progressDiv.textContent =
        "正在上传纯数据 " + (i + 1) + " / " + totalChunks;
      // 判断是否是最后一个分片
      const isLastChunk = i === totalChunks - 1;
      const result = await uploadChunk(
        chunk,
        i,
        totalChunks,
        id,
        isLastChunk,
        "SavePureData",
        template_id
      );
      if (!result.success) {
        progressDiv.textContent = "分片上传失败";
      }
      // 如果是最后一个分片且合并成功
      if (isLastChunk && result.merged) {
        progressDiv.textContent = "数据保存成功";
        await new Promise((resolve) => setTimeout(resolve, 2000));
        removeProgressDiv();
      }
    }
    loadAndMergePureData();
  } catch (error) {
    console.error("纯数据保存失败:", error);
    progressDiv.textContent = "纯数据保存失败: " + error.message;
    await new Promise((resolve) => setTimeout(resolve, 2000));
    removeProgressDiv();
  }
}

// 从服务器获取数据
async function getDataFromServer() {
  const id = (await getUrlParam("id"));
  const template_id = (await getUrlParam("template_id")) || id;

  try {
    // 获取表格总数，优先使用 template_id 查模板
    const countResponse = await fetch(
      `http://${url}:3000/api/getSheetCount?id=${id}&template_id=${template_id}`
    );
    const { count, success } = await countResponse.json();
    if (!success) return {};
    // 循环获取表格
    let sheets = [];
    for (let i = 0; i < count; i++) {
      try {
        document.getElementById("loading-num").textContent =
          "正在加载第" + (i + 1) + "个表格";
        const response = await fetch(
          `http://${url}:3000/api/getSheet?id=${id}&template_id=${template_id}&index=${i}`
        );
        const data = await response.json();
        if (data.success) {
          // 在数组后面添加新表格
          sheets.push(data.sheet);
        }
      } catch (error) {
        console.error("加载第" + i + "个表格失败:", error);
      }
    }
    return sheets;
  } catch (error) {
    console.error("获取数据失败:", error);
    return {};
  }
}

let xs = null;

// 获取纯数据
function getPureData() {
  const allData = xs.getData();
  const pureData = [];
  allData.forEach((sheet) => {
    const pureSheet = {
      name: sheet.name,
      rows: [],
    };

    if (sheet.sheetPreId) {
      pureSheet.sheetId = sheet.sheetPreId;
    }

    // 获取最大行列数B
    let maxRow = 0;
    let maxCol = 0;
    let hasData = false;
    // 第一次遍历找出实际的最大行列数
    Object.entries(sheet.rows).forEach(([rowKey, row]) => {
      if (row && row.cells && Object.keys(row.cells).length > 0) {
        Object.keys(row.cells).forEach((colKey) => {
          const cell = row.cells[colKey];
          // 只处理标记为数据格的单元格或带有id的单元格
          if (cell && (cell.isDataCell === true || cell.id)) {
            hasData = true;
            const rowIndex = parseInt(rowKey);
            const colIndex = parseInt(colKey);
            maxRow = Math.max(maxRow, rowIndex + 1);
            maxCol = Math.max(maxCol, colIndex + 1);
          }
        });
      }
    });
    // 如果表格没有数据，直接跳过
    if (!hasData) {
      return;
    }
    // 遍历每一行
    for (let rowIndex = 0; rowIndex < maxRow; rowIndex++) {
      const row = sheet.rows[rowIndex];
      let rowHasData = false;
      const pureRow = { pxl: rowIndex + 1 }; // pxl从1开始
      // 检查该行是否有数据格
      if (row && row.cells) {
        Object.entries(row.cells).forEach(([colKey, cell]) => {
          if (
            cell &&
            (cell.isDataCell === true || cell.id) &&
            cell.text !== undefined &&
            cell.text !== ""
          ) {
            rowHasData = true;
            const colIndex = parseInt(colKey);
            pureRow[`F${colIndex + 1}`] = cell.text;
            if (cell.id) {
              pureRow[`id_${colIndex + 1}`] = cell.id;
            }
          }
        });
      }
      // 如果行没有数据格，跳过这一行
      if (!rowHasData) {
        continue;
      }
      pureSheet.rows.push(pureRow);
    }
    // 只有当表格有数据时才添加到结果中
    if (pureSheet.rows.length > 0) {
      pureData.push(pureSheet);
    }
  });
  return pureData;
}

// 创建文件输入元素
function createFileInput() {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".json";
  input.style.display = "none";
  document.body.appendChild(input);
  return input;
}

// 纯数据合并函数
async function XSSyncPureData(source, infoDataList = null) {
  try {
    // 处理输入可以是文件或直接的数据对象
    let pureData;
    if (source instanceof File) {
      const fileContent = await source.text();
      pureData = JSON.parse(fileContent);
    } else if (source && typeof source.text === "function") {
      pureData = JSON.parse(source.text());
    } else {
      pureData = source || [];
    }

    // 获取当前所有数据
    const allData = xs.getData();

    // 如果有基于单元格 ID 的数据（由 report_info 提供），将其回填到对应单元格中
    if (infoDataList && infoDataList.length > 0) {
      allData.forEach((sheet, sheetIndex) => {
        let matchedInfo = null;
        if (sheet.sheetPreId) {
          matchedInfo = infoDataList.find(r => r.SheetId === sheet.sheetPreId);
          console.log(matchedInfo);
        } else {
          matchedInfo = infoDataList.find(r => !r.SheetId) || infoDataList[0] || {};
          console.log(matchedInfo);
        }

        if (!matchedInfo) return;

        // 遍历当前表的单元格，如果有设置 ID 且匹配到数据，进行赋值
        if (sheet.rows) {
          Object.entries(sheet.rows).forEach(([rowIndexStr, row]) => {
            if (row && row.cells) {
              Object.entries(row.cells).forEach(([colIndexStr, cell]) => {
                if (cell && cell.id && matchedInfo[cell.id] !== undefined) {
                  const ri = parseInt(rowIndexStr, 10);
                  const ci = parseInt(colIndexStr, 10);
                  // 只有在单元格数据和要填入的数据不一致时才更新
                  if (cell.text !== matchedInfo[cell.id]) {
                    xs.cellText(ri, ci, matchedInfo[cell.id], sheetIndex);
                  }
                }
              });
            }
          });
        }
      });
    }

    // 创建sheet名称到索引的映射，提高查找效率
    const sheetIndexMap = new Map();
    allData.forEach((sheet, index) => {
      sheetIndexMap.set(sheet.name, index);
    });

    // 批量更新所有sheet (基于纯数据结构)
    for (const pureSheet of pureData) {
      const sheetIndex = sheetIndexMap.get(pureSheet.name);
      if (sheetIndex === undefined) continue;

      if (pureSheet.rows && Array.isArray(pureSheet.rows)) {
        // 批量更新的单元格
        for (const row of pureSheet.rows) {
          const rowIndex = row.pxl - 1;
          for (const [key, value] of Object.entries(row)) {
            if (key.startsWith("F") && key !== "pxl") {
              const colIndex = parseInt(key.slice(1)) - 1;
              if (!isNaN(colIndex) && value !== undefined && value !== null) {
                xs.cellText(rowIndex, colIndex, value.toString(), sheetIndex);
              }
            }
          }
        }
      }
    }

    // 最后统一渲染一次
    xs.reRender();
  } catch (error) {
    console.error("合并失败:", error);
    throw error;
  }
}

// 优化表格数据
function optimizeSheetData() {
  const allData = xs.getData();

  allData.forEach((sheet) => {
    // 获取数据列表范围
    const dataListRange = sheet.sheetConfig && sheet.sheetConfig.settings ? sheet.sheetConfig.settings.dataListRange : null;

    // 优化rows - 处理连续空对象
    const rows = sheet.rows;
    let emptyCount = 0;
    let lastValidRow = 0;
    let emptyStartKey = 0;

    // 找出连续5个空行后的位置
    const rowKeys = Object.keys(rows)
      .filter((key) => key !== "len")
      .map(Number)
      .sort((a, b) => a - b);
    
    rowKeys.forEach((key) => {
      const row = rows[key];
      if (row && row.cells) {
        Object.keys(row.cells).forEach((ci) => {
          const cell = row.cells[ci];
          if (cell) {
            // 判断是否需要清空文本：
            // 1. 单元格设置为可编辑
            // 2. 单元格设置为数据格 (isDataCell)
            // 3. 单元格属于数据列表范围
            const isEditable = cell.editable === true;
            const isDataCell = cell.isDataCell === true;
            const isInDataList = dataListRange && key >= dataListRange.sri && key <= dataListRange.eri;

            if (isEditable || isDataCell || isInDataList) {
              cell.text = "";
            }
          }
        });
      }
    });

    for (let i = 0; i < rowKeys.length; i++) {
      const key = rowKeys[i];
      const row = rows[key];

      // 检查是否为空对象（没有cells或cells为空对象）
      const isEmpty = !row || !row.cells || Object.keys(row.cells).length === 0;

      if (isEmpty) {
        if (emptyCount === 0) {
          emptyStartKey = key;
        }
        emptyCount++;

        if (emptyCount >= 5) {
          // 找到了连续5个空对象的位置
          emptyStartKey = key;
          break;
        }
      } else {
        // 不是空对象，重置计数
        emptyCount = 0;
        lastValidRow = key;
      }
    }

    // 如果有连续5个空对象以上，设置新的len
    if (emptyCount >= 5) {
      // 删除第五个空对象后的所有空对象
      rowKeys.forEach((key) => {
        if (
          key > emptyStartKey &&
          (rows[key] === undefined ||
            !rows[key].cells ||
            Object.keys(rows[key].cells).length === 0)
        ) {
          delete rows[key];
        }
      });

      // 设置新的len为第五个空对象的key+1
      rows.len = emptyStartKey + 1;
    } else {
      // 没有连续5个空对象，使用最后一个有效行+5作为len
      rows.len = lastValidRow + 6;
    }

    // 优化cols - 计算需要的列数
    let maxColNeeded = 10; // 默认最小值

    Object.keys(rows).forEach((rowKey) => {
      if (rowKey !== "len" && rows[rowKey] && rows[rowKey].cells) {
        const cells = rows[rowKey].cells;
        if (Object.keys(cells).length > 0) {
          // 获取最后一列的索引
          const lastColIndex = Math.max(...Object.keys(cells).map(Number));
          // 取整（8取10, 14取20）
          const roundedValue = Math.ceil(lastColIndex / 10) * 10;
          maxColNeeded = Math.max(maxColNeeded, roundedValue);
        }
      }
    });

    // 设置cols.len为计算出的最大列数
    sheet.cols.len = maxColNeeded;

    // 更新sheetConfig
    if (sheet.sheetConfig && sheet.sheetConfig.settings) {
      sheet.sheetConfig.settings.row.len = rows.len;
      sheet.sheetConfig.settings.col.len = maxColNeeded;
    } else {
      // 如果sheetConfig不存在，创建它
      sheet.sheetConfig = sheet.sheetConfig || {};
      sheet.sheetConfig.settings = sheet.sheetConfig.settings || {};
      sheet.sheetConfig.settings.row = sheet.sheetConfig.settings.row || {};
      sheet.sheetConfig.settings.col = sheet.sheetConfig.settings.col || {};

      sheet.sheetConfig.settings.row.len = rows.len;
      sheet.sheetConfig.settings.col.len = maxColNeeded;
    }
  });

  return allData;
}

// 在预览和启用模式下加载纯数据并合并
async function loadAndMergePureData() {
  document.getElementById("loading-num").textContent = "正在加载数据";
  const mode = sessionStorage.getItem("mode");

  // 只在预览和启用模式下执行
  if (mode !== "preview" && mode !== "enabled") {
    return;
  }

  try {
    const id = (await getUrlParam("id")) || 0;
    const response = await fetch(
      `http://${url}:3000/api/getPureData?id=${id}`
    );
    const result = await response.json();

    if (result.success) {
      // 使用合并方法将纯数据和信息数据合并到表格
      await XSSyncPureData(
        result.data ? { text: () => JSON.stringify(result.data) } : null,
        result.infoDataList || null
      );
    }

    document.getElementById("loading-num").textContent = "数据加载完成";
  } catch (error) {
    console.error("加载数据失败:", error);
  }
}

// 初始化表格
let isProcessingAction = false;

async function saveReportTemplate(successMsg) {
  try {
    const id = (await getUrlParam("id")) || sessionStorage.getItem("id") || 0;
    const template_id = await getUrlParam("template_id") || sessionStorage.getItem("template_id");
    const chunkSize = 256 * 256;
    const data = optimizeSheetData();
    const jsonData = JSON.stringify(data);
    const totalChunks = Math.ceil(jsonData.length / chunkSize);
    // 显示进度提示
    document.body.appendChild(progressDiv);
    // 上传所有分片
    for (let i = 0; i < totalChunks; i++) {
      const chunk = jsonData.slice(i * chunkSize, (i + 1) * chunkSize);
      progressDiv.textContent = "正在保存报表专属模板 " + (i + 1) + " / " + totalChunks;
      const isLastChunk = i === totalChunks - 1;
      const result = await uploadChunk(
        chunk,
        i,
        totalChunks,
        id,
        isLastChunk,
        "SaveReportTemplate",
        template_id
      );
      if (!result.success) {
        progressDiv.textContent = "模板保存失败";
        await new Promise((resolve) => setTimeout(resolve, 2000));
        break;
      }
      if (isLastChunk && result.merged) {
        progressDiv.textContent = successMsg || "专属模板保存成功";
        if (result.new_template_id) {
          sessionStorage.setItem("template_id", result.new_template_id);
          // 更新 URL 参数以防刷新丢失（如果有使用 URL 记录 template_id 的情况）
          const urlParams = new URLSearchParams(window.location.search);
          if (urlParams.has("template_id")) {
            urlParams.set("template_id", result.new_template_id);
            window.history.replaceState({}, '', `${window.location.pathname}?${urlParams.toString()}`);
          }
        }
        await new Promise((resolve) => setTimeout(resolve, 2000));
        removeProgressDiv();
      }
    }
  } catch (error) {
    console.error("专属模板保存失败:", error);
    progressDiv.textContent = "专属模板保存失败: " + error.message;
    await new Promise((resolve) => setTimeout(resolve, 2000));
    removeProgressDiv();
  }
}

async function handleSaveAction() {
  if (isProcessingAction) return;
  isProcessingAction = true;
  try {
    await savePureDataToServer(getPureData());
  } finally {
    isProcessingAction = false;
  }
}

async function handleVerifyAction() {
  if (isProcessingAction) return;
  isProcessingAction = true;
  console.log("校验按钮被点击");

  try {
    showLoading();
    document.getElementById("loading-num").textContent = "正在保存数据...";
    // 1. 先调用一遍保存
    await savePureDataToServer(getPureData());

    // 2. 发送校验请求
    const id = (await getUrlParam("id")) || sessionStorage.getItem("id") || 0;
    const template_id = await getUrlParam("template_id") || sessionStorage.getItem("template_id");

    document.getElementById("loading-num").textContent = "正在进行数据校验...";
    const response = await fetch(`http://${url}:3000/api/verifyReport?id=${id}&template_id=${template_id}`);
    const result = await response.json();

    hideLoading();
    if (result.success) {
      if (result.valid) {
        alert("校验通过！");
        sessionStorage.setItem("verify_status", "true");
      } else {
        alert("校验不通过:\n\n" + result.errors.join("\n"));
        sessionStorage.setItem("verify_status", "false");
      }
    } else {
      alert("校验请求失败: " + result.message);
      sessionStorage.setItem("verify_status", "false");
    }
  } catch (error) {
    hideLoading();
    console.error("校验请求异常:", error);
    alert("校验请求发生异常");
    sessionStorage.setItem("verify_status", "false");
  } finally {
    isProcessingAction = false;
  }
}

async function handleSubmitAction() {
  if (isProcessingAction) return;

  const verifyStatus = sessionStorage.getItem("verify_status");
  if (!verifyStatus) {
    alert("请先进行校验后再提交！");
    return;
  }
  if (verifyStatus === "false") {
    alert("校验不通过，不允许提交，请修改数据后再重试！");
    return;
  }

  isProcessingAction = true;
  console.log("提交按钮被点击");

  try {
    showLoading();
    document.getElementById("loading-num").textContent = "正在保存数据...";
    // 提交前再保存一次
    await savePureDataToServer(getPureData());

    const id = await getUrlParam("id") || sessionStorage.getItem("id");

    document.getElementById("loading-num").textContent = "正在提交报表...";
    const response = await fetch(`http://${url}:3000/api/submitReport?id=${id}`);
    const result = await response.json();
    hideLoading();

    if (result.success) {
      alert("上报成功！");
      sessionStorage.setItem("isReported", "true");
      sessionStorage.setItem("mode", "preview");

      // 清空表格并重新加载
      document.getElementById("x-spreadsheet-demo").innerHTML = "";
      await load();
    } else {
      alert("上报失败：" + result.message);
    }
  } catch (error) {
    hideLoading();
    console.error("提交异常:", error);
    alert("提交请求发生异常");
  } finally {
    isProcessingAction = false;
  }
}

async function handleWithdrawAction() {
  if (isProcessingAction) return;
  isProcessingAction = true;
  console.log("撤回按钮被点击");

  try {
    const id = await getUrlParam("id") || sessionStorage.getItem("id");
    showLoading();
    document.getElementById("loading-num").textContent = "正在撤回报表...";
    const response = await fetch(`http://${url}:3000/api/withdrawReport?id=${id}`);
    const result = await response.json();
    hideLoading();

    if (result.success) {
      alert("撤回成功！");
      sessionStorage.setItem("isReported", "false");
      sessionStorage.setItem("mode", "enabled");
      sessionStorage.setItem("verify_status", ""); // 撤回后重置校验状态

      // 清空表格并重新加载
      document.getElementById("x-spreadsheet-demo").innerHTML = "";
      await load();
    } else {
      alert("撤回失败：" + result.message);
    }
  } catch (error) {
    hideLoading();
    console.error("撤回异常:", error);
    alert("撤回请求发生异常");
  } finally {
    isProcessingAction = false;
  }
}

async function handleAddRowAction() {
  if (isProcessingAction) return;
  isProcessingAction = true;
  try {
    if (xs && xs.sheet && xs.sheet.contextMenu) {
      xs.sheet.contextMenu.itemClick('insert-row');
      // 保存专属模板并显示“新增成功”
      await saveReportTemplate("新增成功");
    }
  } catch (error) {
    console.error(error);
  } finally {
    isProcessingAction = false;
  }
}

async function handleDeleteRowAction() {
  if (isProcessingAction) return;
  isProcessingAction = true;
  try {
    if (xs && xs.sheet && xs.sheet.contextMenu) {
      xs.sheet.contextMenu.itemClick('delete-row');
      // 保存专属模板并显示“删除成功”
      await saveReportTemplate("删除成功");
    }
  } catch (error) {
    console.error(error);
  } finally {
    isProcessingAction = false;
  }
}

async function load() {
  showLoading();
  var saveIcon =
    "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='rgba(44,44,44,1)'%3E%3Cpath d='M4 3H18L20.7071 5.70711C20.8946 5.89464 21 6.149 21 6.41421V20C21 20.5523 20.5523 21 20 21H4C3.44772 21 3 20.5523 3 20V4C3 3.44772 3.44772 3 4 3ZM7 4V9H16V4H7ZM6 12V19H18V12H6ZM13 5H15V8H13V5Z'%3E%3C/path%3E%3C/svg%3E";
  var pureSaveIcon =
    "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='rgba(44,44,44,1)'%3E%3Cpath d='M4 3H17L20.7071 6.70711C20.8946 6.89464 21 7.149 21 7.41421V20C21 20.5523 20.5523 21 20 21H4C3.44772 21 3 20.5523 3 20V4C3 3.44772 3.44772 3 4 3ZM12 18C13.6569 18 15 16.6569 15 15C15 13.3431 13.6569 12 12 12C10.3431 12 9 13.3431 9 15C9 16.6569 10.3431 18 12 18ZM5 5V9H15V5H5Z'%3E%3C/path%3E%3C/svg%3E";
  var downloadIcon =
    "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='rgba(44,44,44,1)'%3E%3Cpath d='M3 19H21V21H3V19ZM13 9H20L12 17L4 9H11V1H13V9Z'%3E%3C/path%3E%3C/svg%3E";
  var previewEl = document.createElement("img");
  previewEl.src =
    "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBzdGFuZGFsb25lPSJubyI/PjwhRE9DVFlQRSBzdmcgUFVCTElDICItLy9XM0MvL0RURCBTVkcgMS4xLy9FTiIgImh0dHA6Ly93d3cudzMub3JnL0dyYXBoaWNzL1NWRy8xLjEvRFREL3N2ZzExLmR0ZCI+PHN2ZyB0PSIxNjIxMzI4NTkxMjQzIiBjbGFzcz0iaWNvbiIgdmlld0JveD0iMCAwIDEwMjQgMTAyNCIgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHAtaWQ9IjU2NjMiIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB3aWR0aD0iMjAwIiBoZWlnaHQ9IjIwMCI+PGRlZnM+PHN0eWxlIHR5cGU9InRleHQvY3NzIj48L3N0eWxlPjwvZGVmcz48cGF0aCBkPSJNNTEyIDE4Ny45MDRhNDM1LjM5MiA0MzUuMzkyIDAgMCAwLTQxOC41NiAzMTUuNjQ4IDQzNS4zMjggNDM1LjMyOCAwIDAgMCA4MzcuMTIgMEE0MzUuNDU2IDQzNS40NTYgMCAwIDAgNTEyIDE4Ny45MDR6TTUxMiAzMjBhMTkyIDE5MiAwIDEgMSAwIDM4NCAxOTIgMTkyIDAgMCAxIDAtMzg0eiBtMCA3Ni44YTExNS4yIDExNS4yIDAgMSAwIDAgMjMwLjQgMTE1LjIgMTE1LjIgMCAwIDAgMC0yMzAuNHpNMTQuMDggNTAzLjQ4OEwxOC41NiA0ODUuNzZsNC44NjQtMTYuMzg0IDQuOTI4LTE0Ljg0OCA4LjA2NC0yMS41NjggNC4wMzItOS43OTIgNC43MzYtMTAuODggOS4zNDQtMTkuNDU2IDEwLjc1Mi0yMC4wOTYgMTIuNjA4LTIxLjMxMkE1MTEuNjE2IDUxMS42MTYgMCAwIDEgNTEyIDExMS4xMDRhNTExLjQ4OCA1MTEuNDg4IDAgMCAxIDQyNC41MTIgMjI1LjY2NGwxMC4yNCAxNS42OGMxMS45MDQgMTkuMiAyMi41OTIgMzkuMTA0IDMyIDU5Ljc3NmwxMC40OTYgMjQuOTYgNC44NjQgMTMuMTg0IDYuNCAxOC45NDQgNC40MTYgMTQuODQ4IDQuOTkyIDE5LjM5Mi0zLjIgMTIuODY0LTMuNTg0IDEyLjgtNi40IDIwLjA5Ni00LjQ4IDEyLjYwOC00Ljk5MiAxMi45MjhhNTExLjM2IDUxMS4zNiAwIDAgMS0xNy4yOCAzOC40bC0xMi4wMzIgMjIuNC0xMS45NjggMjAuMDk2QTUxMS41NTIgNTExLjU1MiAwIDAgMSA1MTIgODk2YTUxMS40ODggNTExLjQ4OCAwIDAgMS00MjQuNDQ4LTIyNS42bC0xMS4zMjgtMTcuNTM2YTUxMS4yMzIgNTExLjIzMiAwIDAgMS0xOS44NC0zNS4wMDhMNTMuMzc2IDYxMS44NGwtOC42NC0xOC4yNC0xMC4xMTItMjQuMTI4LTcuMTY4LTE5LjY0OC04LjMyLTI2LjYyNC0yLjYyNC05Ljc5Mi0yLjQ5Ni05LjkyeiIgcC1pZD0iNTY2NCI+PC9wYXRoPjwvc3ZnPg==";
  previewEl.width = 16;
  previewEl.height = 16;
  var syncIcon =
    "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='rgba(44,44,44,1)'%3E%3Cpath d='M12 22C6.47715 22 2 17.5228 2 12C2 6.47715 6.47715 2 12 2C17.5228 2 22 6.47715 22 12C22 17.5228 17.5228 22 12 22ZM16.8201 17.0761C18.1628 15.8007 19 13.9981 19 12C19 8.13401 15.866 5 12 5C10.9391 5 9.9334 5.23599 9.03241 5.65834L10.0072 7.41292C10.6177 7.14729 11.2917 7 12 7C14.7614 7 17 9.23858 17 12H14L16.8201 17.0761ZM14.9676 18.3417L13.9928 16.5871C13.3823 16.8527 12.7083 17 12 17C9.23858 17 7 14.7614 7 12H10L7.17993 6.92387C5.83719 8.19929 5 10.0019 5 12C5 15.866 8.13401 19 12 19C13.0609 19 14.0666 18.764 14.9676 18.3417Z'%3E%3C/path%3E%3C/svg%3E";
  var serverData = await getDataFromServer();
  xs = x_spreadsheet("#x-spreadsheet-demo", {
    showToolbar: true,
    showGrid: true,
    showBottomBar: true,
    extendToolbar: {
      left: [
        {
          tip: "保存",
          icon: saveIcon,
          onClick: async (_data, _sheet) => {
            await saveToServer(optimizeSheetData());
          },
        },
        {
          tip: "保存纯数据",
          icon: pureSaveIcon,
          onClick: async (_data, _sheet) => {
            await savePureDataToServer(getPureData());
          },
        },
        {
          tip: "纯数据合并",
          icon: syncIcon,
          onClick: async (_data, _sheet) => {
            const fileInput = createFileInput();
            // 处理文件选择
            fileInput.onchange = async (e) => {
              const file = e.target.files[0];
              if (file) {
                if (file.type !== "application/json") {
                  alert("请选择 JSON 文件");
                  return;
                }
                await XSSyncPureData(file);
              }
              // 清理临时创建的输入元素
              document.body.removeChild(fileInput);
            };
            // 触发文件选择对话框
            fileInput.click();
          },
        },
        {
          tip: "导出Excel",
          icon: downloadIcon,
          onClick: (_data, _sheet) => {
            xs.exportExcel();
          },
        },
      ],
      right: [
        {
          tip: "预览",
          el: previewEl,
          onClick: (data, sheet) => {
            console.log("预览", data, sheet);
          },
        },
      ],
    },
  })
    .loadData(Object.keys(serverData).length > 0 ? serverData : data)
    .change((_cdata) => { });
  xs.initCustomFunctions();
  if (Object.keys(serverData).length === 0) {
    hideLoading();
  }
  xs.on("cell-selected", (_cell, _ri, _ci) => {
    const currentMode = sessionStorage.getItem("mode");
    if (currentMode === "enabled") {
      const dataListRange = xs.sheet.data.rows.getDataListRange();
      const addRowBtn = document.getElementById("add-row-btn");
      const delRowBtn = document.getElementById("delete-row-btn");
      if (dataListRange && _ri >= dataListRange.sri && _ri <= dataListRange.eri) {
        if (addRowBtn) addRowBtn.style.display = "block";
        if (delRowBtn) delRowBtn.style.display = "block";
      } else {
        if (addRowBtn) addRowBtn.style.display = "none";
        if (delRowBtn) delRowBtn.style.display = "none";
      }
    }
  })
    .on("cell-edited", (_text, _ri, _ci) => { })
    .on("pasted-clipboard", (_data) => { })
    .on("grid-load", (_text, _ri, _ci) => {
      console.log("grid-load", xs.sheet.data.getVariables());
    });

  // 获取当前模式
  var mode = sessionStorage.getItem("mode");

  // 防止重复绑定全局事件
  if (!window.hasKeydownListener) {
    document.addEventListener("keydown", async function (e) {
      if ((e.ctrlKey || e.metaKey) && e.key === "s") {
        e.preventDefault();

        // 动态获取最新模式
        const currentMode = sessionStorage.getItem("mode");

        if (currentMode === "design") {
          // 设计模式下保存整个表格
          await saveToServer(optimizeSheetData());
        } else {
          // 预览模式和启用模式下只保存纯数据
          await handleSaveAction();
        }
      }
    });
    window.hasKeydownListener = true;
  }

  if (mode === "preview" || mode === "enabled") {
    // 完整表格加载完成后，加载纯数据并合并
    await loadAndMergePureData();
  }

  // 根据模式控制顶栏显示和绑定事件
  var topBar = document.getElementById("app-top-bar");
  if (topBar) {
    if (mode === "enabled" || mode === "preview") {
      topBar.style.display = "flex";

      // 控制按钮显示隐藏
      const isReported = sessionStorage.getItem("isReported") === "true";
      const saveBtn = document.getElementById("save-pure-data-btn");
      const submitBtn = document.getElementById("submit-btn");
      const verifyBtn = document.getElementById("verify-btn");
      const withdrawBtn = document.getElementById("withdraw-btn");
      const addRowBtn = document.getElementById("add-row-btn");
      const delRowBtn = document.getElementById("delete-row-btn");

      if (isReported) {
        if (saveBtn) saveBtn.style.display = "none";
        if (submitBtn) submitBtn.style.display = "none";
        if (verifyBtn) verifyBtn.style.display = "none";
        if (withdrawBtn) withdrawBtn.style.display = "block";
        if (addRowBtn) addRowBtn.style.display = "none";
        if (delRowBtn) delRowBtn.style.display = "none";
      } else {
        if (saveBtn) saveBtn.style.display = "block";
        if (submitBtn) submitBtn.style.display = "block";
        if (verifyBtn) verifyBtn.style.display = "block";
        if (withdrawBtn) withdrawBtn.style.display = "none";
        // 新增和删除按钮会在 cell-selected 时动态控制
      }

      if (addRowBtn) addRowBtn.onclick = handleAddRowAction;
      if (delRowBtn) delRowBtn.onclick = handleDeleteRowAction;

      // 绑定返回列表按钮事件
      var backBtn = document.getElementById("back-to-list-btn");
      if (backBtn) {
        backBtn.onclick = function () {
          document.getElementById("x-spreadsheet-demo").style.display = "none";
          document.getElementById("app-top-bar").style.display = "none";
          document.getElementById("report-list-container").style.display = "block";
          // 返回列表时重新加载刷新状态
          loadReportList();
        };
      }

      // 绑定保存按钮事件
      if (saveBtn) {
        saveBtn.onclick = handleSaveAction;
      }

      // 绑定导出按钮事件
      var exportBtn = document.getElementById("export-btn");
      if (exportBtn) {
        exportBtn.onclick = function () {
          if (xs) {
            xs.exportExcel();
          }
        };
      }

      // 提交、校验、撤回的点击事件
      if (submitBtn) {
        submitBtn.onclick = handleSubmitAction;
      }

      if (verifyBtn) {
        verifyBtn.onclick = handleVerifyAction;
      }

      if (withdrawBtn) {
        withdrawBtn.onclick = handleWithdrawAction;
      }

      var printBtn = document.getElementById("print-btn");
      if (printBtn) {
        printBtn.onclick = function () {
          console.log("打印按钮被点击");
          if (xs) {
            // 获取所有的表格实例数据对象
            const allDatas = xs.datas;
            xs.sheet.print.preview(allDatas);
          }
        };
      }

    } else {
      topBar.style.display = "none";
    }
  }

  hideLoading();
}
