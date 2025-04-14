// 显示加载动画
function showLoading() {
  document.getElementById("loadingOverlay").style.display = "flex";
}

// 隐藏加载动画
function hideLoading() {
  document.getElementById("loadingOverlay").style.display = "none";
}

// 获取URL参数方法
async function getUrlParam(name) {
  var urlParams = new URLSearchParams(window.location.search);
  return urlParams.get(name);
}

// 上传分片
async function uploadChunk(
  chunk,
  chunkIndex,
  totalChunks,
  id,
  isLastChunk,
  action = "SaveXspreadSheet"
) {
  // 包装数据为对象格式
  const payload = {
    data: chunk,
  };
  const response = await fetch(
    "/Jc_Interface/Result_DataApi.ashx?id=" +
      id +
      "&chunkIndex=" +
      chunkIndex +
      "&totalChunks=" +
      totalChunks +
      "&pure=" +
      isLastChunk +
      "&do=" +
      action,
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
        "SaveXspreadSheet"
      );
      if (!result.success) {
        progressDiv.textContent = "分片上传失败";
      }
      // 如果是最后一个分片且合并成功
      if (isLastChunk && result.merged) {
        progressDiv.textContent = "保存成功";
        await new Promise((resolve) => setTimeout(resolve, 2000));
        document.body.removeChild(progressDiv);
      }
    }
  } catch (error) {
    console.error("保存失败:", error);
    progressDiv.textContent = "保存失败: " + error.message;
    await new Promise((resolve) => setTimeout(resolve, 2000));
    document.body.removeChild(progressDiv);
  }
}

// 保存纯数据到服务器
async function savePureDataToServer(data) {
  try {
    const id = (await getUrlParam("id")) || 0;
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
        "SavePureData"
      );
      if (!result.success) {
        progressDiv.textContent = "分片上传失败";
      }
      // 如果是最后一个分片且合并成功
      if (isLastChunk && result.merged) {
        progressDiv.textContent = "纯数据保存成功";
        await new Promise((resolve) => setTimeout(resolve, 2000));
        document.body.removeChild(progressDiv);
      }
    }
  } catch (error) {
    console.error("纯数据保存失败:", error);
    progressDiv.textContent = "纯数据保存失败: " + error.message;
    await new Promise((resolve) => setTimeout(resolve, 2000));
    document.body.removeChild(progressDiv);
  }
}

// 从服务器获取数据
async function getDataFromServer() {
  const id = (await getUrlParam("id")) || 0;
  try {
    // 获取表格总数
    const countResponse = await fetch(
      "/Jc_Interface/Result_DataApi.ashx?id=" +
        id +
        "&chunkIndex=0&do=GetXspreadSheet"
    );
    const { count, success } = await countResponse.json();

    if (!success) return {};
    // 循环获取表格
    let sheets = [];
    for (let i = 1; i <= count; i++) {
      try {
        document.getElementById("loading-num").textContent =
          "正在加载第" + i + "个表格";
        const response = await fetch(
          "/Jc_Interface/Result_DataApi.ashx?id=" +
            id +
            "&chunkIndex=" +
            i +
            "&do=GetXspreadSheet"
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
    hideLoading();
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
    // 获取最大行列数B
    let maxRow = 0;
    let maxCol = 0;
    let hasData = false;
    // 第一次遍历找出实际的最大行列数
    Object.entries(sheet.rows).forEach(([rowKey, row]) => {
      if (row && row.cells && Object.keys(row.cells).length > 0) {
        Object.keys(row.cells).forEach((colKey) => {
          const cell = row.cells[colKey];
          // 只处理标记为数据格的单元格
          if (cell && cell.isDataCell === true) {
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
            cell.isDataCell === true &&
            cell.text !== undefined &&
            cell.text !== ""
          ) {
            rowHasData = true;
            const colIndex = parseInt(colKey);
            pureRow[`F${colIndex + 1}`] = cell.text;
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
async function XSSyncPureData(source) {
  document.body.appendChild(progressDiv);
  progressDiv.textContent = "正在处理数据...";
  try {
    // 处理输入可以是文件或直接的数据对象
    let pureData;
    if (source instanceof File) {
      // 读取上传的JSON文件
      progressDiv.textContent = "正在读取文件...";
      const fileContent = await source.text();
      pureData = JSON.parse(fileContent);
    } else if (typeof source.text === "function") {
      // 从API返回的数据
      pureData = JSON.parse(source.text());
    } else {
      // 直接是数据对象
      pureData = source;
    }

    // 获取当前所有数据
    const allData = xs.getData();

    // 遍历每个sheet
    for (const pureSheet of pureData) {
      // 在当前数据中查找对应名称的sheet
      const sheetIndex = allData.findIndex(
        (sheet) => sheet.name === pureSheet.name
      );
      if (sheetIndex !== -1) {
        updatedSheets++;
        progressDiv.textContent = `正在更新表格 "${pureSheet.name}"...`;

        // 处理每个表格的行数据
        if (pureSheet.rows && Array.isArray(pureSheet.rows)) {
          updatedCells += await processBatch(pureSheet, sheetIndex);
        }

        // 每个表格更新完后重新渲染
        xs.reRender();
      } else {
        console.warn(`未找到名为 "${pureSheet.name}" 的表格`);
      }
    }

    progressDiv.textContent = `更新完成，共更新了 ${updatedSheets} 个表格，${updatedCells} 个单元格`;
    await new Promise((resolve) => setTimeout(resolve, 2000));
    document.body.removeChild(progressDiv);
  } catch (error) {
    console.error("合并失败:", error);
    progressDiv.textContent = "合并失败: " + error.message;
    await new Promise((resolve) => setTimeout(resolve, 2000));
    document.body.removeChild(progressDiv);
  }
}

// 处理每个表格的行数据，分批处理
async function processBatch(pureSheet, sheetIndex) {
  const batchSize = 100; // 每批处理的行数
  const rows = pureSheet.rows || [];
  let updatedCells = 0;

  for (let i = 0; i < rows.length; i += batchSize) {
    const batch = rows.slice(i, i + batchSize);
    for (const row of batch) {
      updatedCells += await processRow(row, sheetIndex);
    }
    // 让事件循环有机会处理其他任务
    await new Promise((resolve) => setTimeout(resolve, 0));
  }

  return updatedCells;
}

// 处理单行数据
async function processRow(row, sheetIndex) {
  let updatedCells = 0;
  const rowIndex = row.pxl - 1; // pxl从1开始，需要减1来匹配表格索引

  for (const [key, value] of Object.entries(row)) {
    if (key.startsWith("F") && key !== "pxl") {
      const colIndex = parseInt(key.slice(1)) - 1; // 从F后面提取列号，并减1作为列索引
      // 确保列号是有效的数字且有数据
      if (!isNaN(colIndex) && value !== undefined && value !== null) {
        // 将数字转换为字符串
        const cellValue = value.toString();
        // 使用cellText更新单元格
        const currentText = xs.cell(rowIndex, colIndex, sheetIndex)?.text;
        if (currentText !== cellValue) {
          xs.cellText(rowIndex, colIndex, cellValue, sheetIndex);
          updatedCells++;
        }
      }
    }
  }

  return updatedCells;
}

// 优化表格数据
function optimizeSheetData() {
  const allData = xs.getData();

  allData.forEach((sheet) => {
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
  const urlParams = new URLSearchParams(window.location.search);
  const mode = urlParams.get("mode");

  // 只在预览和启用模式下执行
  if (mode !== "preview" && mode !== "enabled") {
    return;
  }

  try {
    document.body.appendChild(progressDiv);

    const id = (await getUrlParam("id")) || 0;
    const response = await fetch(
      "/Jc_Interface/Result_DataApi.ashx?id=" + id + "&do=GetPureData"
    );
    const result = await response.json();

    if (result.success && result.data) {
      // 使用合并方法将纯数据合并到表格
      await XSSyncPureData({
        text: function () {
          return JSON.stringify(result.data);
        },
      });
    }
  } catch (error) {
    console.error("加载数据失败:", error);
  }
}

// 初始化表格
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
    .change((_cdata) => {});
  xs.initCustomFunctions();
  if (Object.keys(serverData).length === 0) {
    hideLoading();
  }
  xs.on("cell-selected", (_cell, _ri, _ci) => {})
    .on("cell-edited", (_text, _ri, _ci) => {})
    .on("pasted-clipboard", (_data) => {})
    .on("grid-load", (_text, _ri, _ci) => {
      console.log("grid-load", xs.sheet.data.getVariables());
    });

  // 获取当前模式
  var urlParams = new URLSearchParams(window.location.search);
  var mode = urlParams.get("mode");

  document.addEventListener("keydown", async function (e) {
    if ((e.ctrlKey || e.metaKey) && e.key === "s") {
      e.preventDefault();

      if (mode === "design") {
        // 设计模式下保存整个表格
        await saveToServer(optimizeSheetData());
      } else {
        // 预览模式和启用模式下只保存纯数据
        await savePureDataToServer(getPureData());
      }
    }
  });

  if (mode === "preview" || mode === "enabled") {
    // 完整表格加载完成后，加载纯数据并合并
    await loadAndMergePureData();
  }
}

// 加载并合并纯数据
async function loadAndMergePureData() {
  try {
    const pureData = await getPureDataFromServer();
    if (pureData && Object.keys(pureData).length > 0) {
      await XSSyncPureData(null, pureData);
    }
  } catch (error) {
    console.error("加载纯数据失败", error);
  }
}
