import ExcelJS from "exceljs";
import tinycolor from "tinycolor2";

// 添加中文字体映射
const FONT_MAP = {
  宋体: "SimSun",
  黑体: "SimHei",
  微软雅黑: "Microsoft YaHei",
  楷体: "KaiTi",
  仿宋: "FangSong",
  新宋体: "NSimSun",
  等线: "DengXian",
  华文楷体: "STKaiti",
  华文宋体: "STSong",
  华文仿宋: "STFangsong",
  华文中宋: "STZhongsong",
  华文黑体: "STHeiti",
  方正舒体: "FZShuTi",
  方正姚体: "FZYaoti",
};

/**
 * 生成时间字符串作为文件名
 */
function generateFileName() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const day = now.getDate();
  const hour = String(now.getHours()).padStart(2, "0");
  const minute = String(now.getMinutes()).padStart(2, "0");
  const second = String(now.getSeconds()).padStart(2, "0");

  return `${year}年${month}月${day}日${hour}：${minute}：${second}.xlsx`;
}

/**
 * 将颜色转换为Excel格式的ARGB
 */
function convertToARGB(color) {
  if (!color) return undefined;
  const rgb = tinycolor(color).toRgb();
  const rHex = parseInt(rgb.r).toString(16).padStart(2, "0");
  const gHex = parseInt(rgb.g).toString(16).padStart(2, "0");
  const bHex = parseInt(rgb.b).toString(16).padStart(2, "0");
  const aHex = "FF";
  return aHex + rHex + gHex + bHex;
}

/**
 * 处理单元格样式
 */
function applyCellStyle(excelCell, cellStyle) {
  if (!cellStyle) return;

  // 对齐方式
  const alignment = {};
  if (cellStyle.align) alignment.horizontal = cellStyle.align;
  if (cellStyle.valign) alignment.vertical = cellStyle.valign;
  if (cellStyle.textwrap) alignment.wrapText = true;
  if (Object.keys(alignment).length > 0) {
    excelCell.alignment = alignment;
  }

  // 背景色
  if (cellStyle.bgcolor) {
    excelCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: convertToARGB(cellStyle.bgcolor) },
    };
  }

  // 字体设置
  const fontOptions = {
    name: FONT_MAP[cellStyle.font?.name] || cellStyle.font?.name || "Arial",
    size: cellStyle.font?.size || 10,
    bold: cellStyle.font?.bold || false,
    italic: cellStyle.font?.italic || false,
  };

  // 字体颜色
  if (cellStyle.color) {
    fontOptions.color = { argb: convertToARGB(cellStyle.color) };
  }

  // 文本装饰 - 检查所有可能的来源
  const textDecoration =
    cellStyle.textDecoration || cellStyle.font?.textDecoration;
  if (textDecoration) {
    // 设置下划线
    if (textDecoration.includes("underline")) {
      fontOptions.underline = "single";
    }
    // 设置删除线
    if (textDecoration.includes("line-through")) {
      fontOptions.strike = true;
    }
  }

  // 应用字体设置
  excelCell.font = fontOptions;

  // 边框样式
  if (cellStyle.border) {
    const borderStyle = {};
    ["top", "right", "bottom", "left"].forEach((side) => {
      if (cellStyle.border[side]) {
        const [style, color] = cellStyle.border[side];
        borderStyle[side] = {
          style: style || "thin",
          color: { argb: convertToARGB(color) },
        };
      }
    });
    if (Object.keys(borderStyle).length > 0) {
      excelCell.border = borderStyle;
    }
  }
}

/**
 * 处理单元格合并
 */
function getMergeRange(cell, rowIndex, colIndex) {
  if (!cell || !cell.merge) return null;
  const [rowSpan, colSpan] = cell.merge;
  // 计算起始和结束位置
  const startRow = parseInt(rowIndex) + 1;
  const startCol = parseInt(colIndex) + 1;
  const endRow = startRow + rowSpan;
  const endCol = startCol + colSpan;
  return {
    startCell: { row: startRow, col: startCol },
    endCell: { row: endRow, col: endCol },
  };
}

// 添加辅助函数 - 将列号转换为Excel列字母
function getExcelColName(n) {
  let result = "";
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

/**
 * 导出Excel文件
 */
const ExcelExport = async function (datas) {
  try {
    const workbook = new ExcelJS.Workbook();
    workbook.created = new Date();
    workbook.modified = new Date();
    workbook.creator = "x-spreadsheet";
    workbook.lastModifiedBy = "x-spreadsheet";

    datas.forEach((sheet) => {
      const worksheet = workbook.addWorksheet(sheet.name);

      // 设置列宽
      if (sheet.cols) {
        Object.entries(sheet.cols).forEach(([colIndex, col]) => {
          if (col && col.width) {
            // Excel列宽与像素的转换比例约为7.5
            worksheet.getColumn(parseInt(colIndex) + 1).width = col.width / 7.5;
          }
        });
      }

      // 处理行数据
      if (sheet.rows) {
        Object.entries(sheet.rows).forEach(([rowIndex, row]) => {
          if (!row || !row.cells) return;

          const excelRow = worksheet.getRow(parseInt(rowIndex) + 1);

          // 设置行高
          if (row.height) {
            // Excel行高与像素的转换比例约为0.75
            excelRow.height = row.height * 0.75;
          }

          // 处理单元格
          Object.entries(row.cells).forEach(([colIndex, cell]) => {
            if (!cell) return;

            // 处理合并单元格
            if (cell.merge) {
              const mergeInfo = getMergeRange(cell, rowIndex, colIndex);
              if (mergeInfo) {
                const { startCell, endCell } = mergeInfo;
                const cellValue = cell.text || "";
                // 先执行合并
                worksheet.mergeCells(
                  startCell.row,
                  startCell.col,
                  endCell.row,
                  endCell.col
                );
                // 获取Excel格式的单元格引用并设置值和样式
                const cellRef = `${getExcelColName(startCell.col)}${startCell.row};`;
                const targetCell = worksheet.getCell(cellRef);
                // 确保正确设置值
                if (cellValue.startsWith("=")) {
                  targetCell.value = { formula: cellValue.substring(1) };
                } else {
                  const numValue = Number(cellValue);
                  targetCell.value = isNaN(numValue) ? cellValue : numValue;
                }
                // 设置样式
                targetCell.alignment = {
                  vertical: "middle",
                  horizontal: "center",
                };
                if (cell.style !== undefined && sheet.styles) {
                  applyCellStyle(targetCell, sheet.styles[cell.style]);
                }
              }
            } else {
              const excelCell = excelRow.getCell(parseInt(colIndex) + 1);

              // 设置值
              if (cell.text !== undefined) {
                if (cell.text.startsWith("=")) {
                  excelCell.value = { formula: cell.text.substring(1) };
                } else {
                  // 处理普通文本
                  if (
                    cell.text === "" ||
                    cell.text === null ||
                    cell.text === undefined
                  ) {
                    excelCell.value = null;
                  } else {
                    const numValue = Number(cell.text);
                    excelCell.value = isNaN(numValue) ? cell.text : numValue;
                  }
                }
              }

              // 设置样式
              if (cell.style !== undefined && sheet.styles) {
                applyCellStyle(excelCell, sheet.styles[cell.style]);
              }
            }
          });
        });
      }

      // 设置网格线
      if (sheet.sheetConfig?.gridLine === false) {
        worksheet.views = [{ showGridLines: false }];
      }
    });

    // 生成并下载文件
    const buffer = await workbook.xlsx.writeBuffer();
    const fileName = generateFileName();

    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
  } catch (error) {
    console.error("导出Excel失败:", error);
    throw error;
  }
};

export default ExcelExport;
