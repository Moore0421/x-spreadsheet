import ExcelJs from "exceljs";
import * as XLSX from "xlsx";
import {
  PX_TO_PT,
  CELL_REF_REPLACE_REGEX,
  SHEET_TO_CELL_REF_REGEX,
  INDEXED_COLORS,
  EXTRACT_MSO_NUMBER_FORMAT_REGEX,
} from "./constants";
import { fonts } from "./core/font";
import { npx } from "./canvas/draw";

const avialableFonts = Object.keys(fonts());

const FONT_MAP = {
  宋体: "SimSun",
  黑体: "SimHei",
  微软雅黑: "Microsoft YaHei",
  楷体: "KaiTi",
  仿宋: "FangSong",
  新宋体: "NSimSun",
  华文宋体: "STSong",
  华文仿宋: "STFangsong",
  华文中宋: "STZhongsong",
  华文楷体: "STKaiti",
  华文黑体: "STHeiti",
};

// 添加边框样式映射常量
const BORDER_STYLE_MAP = {
  solid: "thin",
  dashed: "dashed",
  dotted: "dotted",
  double: "double",
  groove: "medium",
  ridge: "thick",
  inset: "thin",
  outset: "thin",
};

// 添加Excel到CSS的反向映射
const EXCEL_TO_CSS_BORDER = {
  thin: "solid",
  medium: "solid",
  thick: "solid",
  dashed: "dashed",
  dotted: "dotted",
  double: "double",
};

const getStylingForClass = (styleTag, className) => {
  const cssRules = styleTag?.sheet?.cssRules || styleTag?.sheet?.rules;
  for (let i = 0; i < cssRules?.length; i++) {
    const cssRule = cssRules[i];
    if (cssRule.selectorText === `.${className}`) {
      return cssRule.style.cssText;
    }
  }
  return "";
};

const parseCssToXDataStyles = (styleString, cellType) => {
  const parsedStyles = {};
  // By default numbers are right aligned.
  if (cellType === "n") parsedStyles.align = "right";
  if (styleString) {
    const fontStyles = {};
    let borderStyles = {};
    const dimensions = {};
    const styles = styleString.split(";");
    const stylesObject = {};
    styles.forEach((style) => {
      const [property, value] = style.split(":");
      if (property && value) stylesObject[property.trim()] = value.trim();
    });

    let gridStatus = false;
    const parsedStylesObject = parseBorderProperties(stylesObject);
    Object.entries(parsedStylesObject).forEach(([property, value]) => {
      switch (property) {
        case "background":
        case "background-color":
          parsedStyles["bgcolor"] = value;
          break;
        case "color":
          parsedStyles["color"] = value;
          break;
        case "text-decoration":
          if (value === "underline") parsedStyles["underline"] = true;
          else if (value === "line-through") parsedStyles["strike"] = true;
          break;
        case "text-align":
          parsedStyles["align"] = value;
          break;
        case "vertical-align":
          parsedStyles["valign"] = ["justify", "center"].includes(value)
            ? "middle"
            : value;
          break;
        case "font-weight":
          const parsedIntValue = parseInt(value);
          fontStyles["bold"] =
            (parsedIntValue !== NaN && parsedIntValue > 400) ||
            value === "bold";
          break;
        case "font-size":
          fontStyles["size"] = parsePtOrPxValue(value, property);
          break;
        case "font-style":
          fontStyles["italic"] = value === "italic";
          break;
        case "font-family":
          const fontName = value.split(",")[0].trim().replace(/['"]/g, "");
          const mappedFont = FONT_MAP[fontName] || fontName;
          const appliedFont =
            avialableFonts.find((font) =>
              mappedFont.toLowerCase().includes(font.toLowerCase())
            ) ?? "Arial";
          fontStyles["name"] = appliedFont;
          break;
        case "border":
        case "border-top":
        case "border-bottom":
        case "border-left":
        case "border-right":
          if (property === "border" && !gridStatus && value === "0px") {
            gridStatus = true;
            break;
          }
          const values = value.split(" ");
          if (values.length >= 3) {
            const width = values[0];
            const style = values[1];
            const color = values.slice(2).join(" ");

            const lineStyle =
              BORDER_STYLE_MAP[style] ||
              (parsePtOrPxValue(width) <= 1
                ? "thin"
                : parsePtOrPxValue(width) <= 2
                  ? "medium"
                  : "thick");

            const parsedValues = [lineStyle, color];

            if (property === "border") {
              borderStyles = {
                top: parsedValues,
                bottom: parsedValues,
                left: parsedValues,
                right: parsedValues,
              };
            } else {
              const side = property.split("-")[1];
              borderStyles[side] = parsedValues;
            }
          }
          break;
        case "width":
          const widthValue = parsePtOrPxValue(value);
          if (widthValue) dimensions.width = widthValue;
          break;
        case "height":
          const heightValue = parsePtOrPxValue(value);
          if (heightValue) dimensions.height = heightValue;
          break;
        case "text-wrap":
          parsedStyles["textwrap"] = value === "wrap";
          break;
      }
    });
    parsedStyles["dimensions"] = dimensions;
    parsedStyles["font"] = fontStyles;
    if (Object.keys(borderStyles).length) parsedStyles["border"] = borderStyles;
    return { parsedStyles, sheetConfig: { gridLine: gridStatus } };
  }
  return { parsedStyles, sheetConfig: { gridLine: false } };
};

const parseBorderProperties = (styles) => {
  const border = {
    "border-top": {},
    "border-right": {},
    "border-bottom": {},
    "border-left": {},
  };
  const others = {};
  const parsedBorders = {};

  // 合并相同边的样式
  const mergeBorderStyles = (side) => {
    const value = border[`border-${side}`];
    if (value.width || value.style || value.color) {
      const width = value.width || "1px";
      const style = value.style || "solid";
      const color = value.color || "#000000";
      return `${width} ${style} ${color}`;
    }
    return null;
  };

  for (const key in styles) {
    if (styles.hasOwnProperty(key)) {
      const parts = key.split("-");
      // 处理完整的border属性
      if (key === "border") {
        const values = styles[key].split(" ");
        if (values.length >= 3) {
          const [width, style, color] = values;
          ["top", "right", "bottom", "left"].forEach((side) => {
            border[`border-${side}`] = { width, style, color };
          });
        }
        continue;
      }

      // 处理单边属性
      if (parts[0] === "border" && parts.length === 2) {
        const side = parts[1];
        if (border[`border-${side}`]) {
          const values = styles[key].split(" ");
          if (values.length >= 3) {
            const [width, style, color] = values;
            border[`border-${side}`] = { width, style, color };
          }
        }
      }
      // 处理分开的边框属性
      else if (parts.length === 3 && parts[0] === "border") {
        const side = parts[1];
        const prop = parts[2];
        if (!border[`border-${side}`]) {
          border[`border-${side}`] = {};
        }
        border[`border-${side}`][prop] = styles[key];
      } else {
        others[key] = styles[key];
      }
    }
  }

  // 处理每个边的样式
  ["top", "right", "bottom", "left"].forEach((side) => {
    const borderStyle = mergeBorderStyles(side);
    if (borderStyle) {
      parsedBorders[`border-${side}`] = borderStyle;
    }
  });

  return { ...parsedBorders, ...others };
};

const parsePtOrPxValue = (value, property = null) => {
  let parsedValue = value;
  if (value) {
    if (value.includes("px")) {
      parsedValue = Math.ceil(Number(value.split("px")[0]));
    } else if (value.includes("pt")) {
      if (property === "font-size")
        parsedValue = Math.ceil(Number(value.split("pt")[0]));
      else parsedValue = Math.ceil(Number(value.split("pt")[0]) / PX_TO_PT);
    }
  }
  return parsedValue;
};

const parseHtmlToText = (function () {
  const entities = [
    ["nbsp", ""],
    ["middot", "·"],
    ["quot", '"'],
    ["apos", "'"],
    ["gt", ">"],
    ["lt", "<"],
    ["amp", "&"],
  ].map(function (x) {
    return [new RegExp("&" + x[0] + ";", "ig"), x[1]];
  });
  return function parseHtmlToText(str) {
    let o = str
      // Remove new lines and spaces from start of content
      .replace(/^[\t\n\r ]+/, "")
      // Remove new lines and spaces from end of content
      .replace(/[\t\n\r ]+$/, "")
      // Added line which removes any white space characters after and before html tags
      .replace(/>\s+/g, ">")
      .replace(/\s+</g, "<")
      // Replace remaining new lines and spaces with space
      .replace(/[\t\n\r ]+/g, " ")
      // Replace <br> tags with new lines
      .replace(/<\s*[bB][rR]\s*\/?>/g, "\n")
      // Strip HTML elements
      .replace(/<[^>]*>/g, "");
    for (let i = 0; i < entities.length; ++i)
      o = o.replace(entities[i][0], entities[i][1]);
    return o;
  };
})();

const generateUniqueId = () => {
  const timestamp = Date.now().toString(36);
  const random = Math.random().toString(36);
  return timestamp + random;
};

// Function to match the pattern and replace it
const replaceCellRefWithNew = (str, getNewCell, opts) => {
  const { isSameSheet, sheetName } = opts;
  const newStr = str.replace(
    isSameSheet ? CELL_REF_REPLACE_REGEX : SHEET_TO_CELL_REF_REGEX,
    (word) => {
      const parts = word.split("!");
      if (parts.length > 1) {
        if (parts[0].replaceAll("'", "") === sheetName) {
          const newCell = getNewCell(parts[1]);
          return `${parts[0]}!${newCell}`;
        } else {
          return word;
        }
      } else if (isSameSheet) {
        const newCell = getNewCell(parts[0]);
        return newCell;
      }
    }
  );
  return newStr;
};

const readExcelFile = (file) => {
  const ExcelWorkbook = new ExcelJs.Workbook();
  const styles = {};
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      ExcelWorkbook.xlsx.load(reader.result).then((workbookIns) => {
        workbookIns.eachSheet((sheet) => {
          const sheetName = sheet?.name;
          styles[sheetName] = {};
          sheet.eachRow({ includeEmpty: true }, (row) => {
            row?.eachCell({ includeEmpty: true }, (cell) => {
              const style = cell.style;
              const address = cell.address;
              styles[sheetName][address] = style;
            });
          });
        });

        const data = new Uint8Array(e.target?.result);
        const wb = XLSX?.read(data, {
          type: "array",
          cellStyles: true,
          sheetStubs: true,
        });

        const workbook = addStylesToWorkbook(styles, wb);
        resolve(workbook);
      });
    };
    reader.onerror = (error) => {
      reject(error);
    };
    reader.readAsArrayBuffer(file);
  });
};

const parseExcelStyleToHTML = (styling, theme) => {
  let styleString = "";
  const parsedStyles = {};
  Object.keys(styling)?.forEach((styleOn) => {
    const style = styling[styleOn];
    switch (styleOn) {
      case "alignment":
        Object.keys(style).forEach((property) => {
          const value = style[property];
          switch (property) {
            case "vertical":
              parsedStyles["display"] = "table-cell";
              parsedStyles["vertical-align"] = value;
              break;
            case "horizontal":
              parsedStyles["text-align"] = value;
              break;
            case "wrapText":
              parsedStyles["text-wrap"] = value ? "wrap" : "nowrap";
              break;
          }
        });
        break;
      case "border":
        Object.keys(style).forEach((side) => {
          const value = style[side];
          switch (side) {
            case "top":
            case "bottom":
            case "right":
            case "left":
              if (value?.style) {
                const borderStyle = EXCEL_TO_CSS_BORDER[value.style] || "solid";
                const borderColor =
                  value.rgb ||
                  (value?.color?.indexed
                    ? INDEXED_COLORS[value.color.indexed]
                    : "#000000");
                parsedStyles[`border-${side}`] =
                  `1px ${borderStyle} ${borderColor}`;
              }
              break;
          }
        });
        break;
      case "fill":
        Object.keys(style)?.forEach((property) => {
          const value = style[property];
          switch (property) {
            case "bgColor":
            case "fgColor":
              if (value?.rgb) {
                parsedStyles["background-color"] = value.rgb.startsWith("#")
                  ? value.rgb
                  : `#${value.rgb}`;
              } else if (value?.argb) {
                parsedStyles["background-color"] = value.argb.startsWith("#")
                  ? `#${value.argb.slice(3)}`
                  : `#${value.argb.slice(2)}`;
              } else if (value?.theme && Object.hasOwn(theme, value.theme))
                parsedStyles["background-color"] =
                  `#${theme[value.theme].rgb}` ?? "#ffffff";
              else if (value?.color?.indexed) {
                const color = INDEXED_COLORS[value.color.indexed] ?? "#ffffff";
                parsedStyles[`background-color`] = color;
              }
              break;
          }
        });
        break;
      case "font":
        Object.keys(style)?.forEach((property) => {
          const value = style[property];
          switch (property) {
            case "bold":
              parsedStyles["font-weight"] = value ? "bold" : "normal";
              break;
            case "color":
              if (value?.rgb) {
                parsedStyles["color"] = value.rgb.startsWith("#")
                  ? value.rgb
                  : `#${value.rgb}`;
              } else if (value?.argb) {
                parsedStyles["color"] = value.argb.startsWith("#")
                  ? `#${value.argb.slice(3)}`
                  : `#${value.argb.slice(2)}`;
              } else if (value?.theme && Object.hasOwn(theme, value.theme)) {
                parsedStyles["color"] =
                  `#${theme[value.theme].rgb}` ?? "#000000";
              } else if (value?.color?.indexed) {
                const color = INDEXED_COLORS[value.color.indexed] ?? "#000000";
                parsedStyles[`color`] = color;
              }
              break;
            case "sz":
            case "size":
              const convertedValue = Number(value) ?? 11;
              parsedStyles["font-size"] = `${convertedValue}px`;
              break;
            case "italic":
              parsedStyles["font-style"] = value ? "italic" : "normal";
              break;
            case "name":
              parsedStyles["font-family"] = value;
              break;
            case "underline":
            case "strike":
              parsedStyles["text-decoration"] = value
                ? property === "underline"
                  ? "underline"
                  : "line-through"
                : "none";
              break;
          }
        });
        break;
      case "dimensions":
        Object.keys(style)?.forEach((property) => {
          const value = style[property];
          switch (property) {
            case "height":
            case "width":
              parsedStyles[property] = value;
              break;
          }
        });
    }
  });

  Object.entries(parsedStyles).forEach(([property, value]) => {
    styleString = `${styleString}${property}:${value};`;
  });

  return styleString;
};

const addStylesToWorkbook = (styles, workbook) => {
  const wb = { ...workbook };
  wb.SheetNames.forEach((sheetName) => {
    const worksheet = wb.Sheets[sheetName];
    if (Object.hasOwn(styles, sheetName)) {
      Object.entries(styles[sheetName]).forEach(([cellAddress, cellStyle]) => {
        if (Object.hasOwn(worksheet, cellAddress)) {
          const { r, c } = XLSX.utils.decode_cell(cellAddress);
          const dimensions = {};

          const height = worksheet?.["!rows"]?.[r]?.hpt;
          const width = worksheet?.["!cols"]?.[c]?.wpx;

          if (height) dimensions.height = `${height / 0.75}px`;
          if (width) dimensions.width = `${width}px`;

          const cellStylesWithDimensions = {
            ...(cellStyle ?? {}),
            dimensions,
          };

          worksheet[cellAddress] = {
            ...worksheet[cellAddress],
            s: parseExcelStyleToHTML(
              cellStylesWithDimensions,
              wb.Themes?.themeElements?.clrScheme ?? {}
            ),
          };
        }
      });
    }
  });
  return wb;
};

const stox = (wb) => {
  const out = [];
  wb.SheetNames.forEach(function (name) {
    const o = { name: name, rows: {}, cols: {}, styles: [] };
    const ws = wb.Sheets[name];
    let gridStatus = false;
    if (!ws || !ws["!ref"]) return;
    const range = XLSX.utils.decode_range(ws["!ref"]);
    // sheet_to_json will lost empty row and col at begin as default

    // Populating 100 rows and a-z columns by default.
    if (range?.e) {
      if (range.e.r < 99) range.e.r = 99;
      if (range.e.c < 25) range.e.c = 25;
    } else {
      range.e = {
        r: 99,
        c: 25,
      };
    }

    range.s = { r: 0, c: 0 };
    const aoa = XLSX.utils.sheet_to_json(ws, {
      raw: false,
      header: 1,
      range: range,
    });

    // 预处理所有单元格样式
    const processedStyles = {};
    Object.keys(ws).forEach((cellRef) => {
      if (cellRef[0] !== "!") {
        // 跳过特殊属性
        const cell = ws[cellRef];
        if (cell.s) {
          processedStyles[cellRef] = {
            style: cell.s,
            type: cell.t,
          };
        }
      }
    });

    aoa.forEach(function (r, i) {
      const cells = {};
      let rowHeight = null;
      r.forEach(function (c, j) {
        const cellRef = XLSX.utils.encode_cell({ r: i, c: j });
        const cell = ws[cellRef] || {};

        // 初始化单元格
        cells[j] = { text: c || "" };

        // 检查是否有预处理的样式
        const processedStyle = processedStyles[cellRef];
        if (processedStyle) {
          const parsedData = parseCssToXDataStyles(
            processedStyle.style,
            processedStyle.type
          );
          const parsedCellStyles = parsedData.parsedStyles;
          const sheetConfig = parsedData.sheetConfig;

          if (!gridStatus && sheetConfig?.gridLine) {
            gridStatus = true;
          }

          // 处理尺寸
          const dimensions = parsedCellStyles.dimensions;
          delete parsedCellStyles.dimensions;

          if (Object.keys(parsedCellStyles).length) {
            const length = o.styles.push(parsedCellStyles);
            cells[j].style = length - 1;
          }

          if (dimensions?.height) rowHeight = dimensions.height;
          if (dimensions?.width) {
            o.cols[j] = { width: dimensions.width };
          }
        }

        // 其他处理
        if (cell.w !== undefined) {
          cells[j].formattedText = cell.w;
        }
        if (cell?.f) {
          cells[j].text = "=" + cell.f;
        }
        if (cell.metadata) {
          cells[j].cellMeta = cell.metadata;
        }
      });

      // 为每列添加空单元格
      for (let j = 0; j <= range.e.c; j++) {
        if (!cells[j]) {
          const cellRef = XLSX.utils.encode_cell({ r: i, c: j });
          const processedStyle = processedStyles[cellRef];
          if (processedStyle) {
            const parsedData = parseCssToXDataStyles(
              processedStyle.style,
              processedStyle.type
            );
            const parsedCellStyles = parsedData.parsedStyles;
            delete parsedCellStyles.dimensions;
            if (Object.keys(parsedCellStyles).length) {
              cells[j] = { text: "" };
              const length = o.styles.push(parsedCellStyles);
              cells[j].style = length - 1;
            }
          }
        }
      }

      if (rowHeight) o.rows[i] = { cells: cells, height: rowHeight };
      else o.rows[i] = { cells: cells };
    });

    o.merges = [];
    (ws["!merges"] || []).forEach(function (merge, i) {
      //Needed to support merged cells with empty content
      if (o.rows[merge.s.r] == null) {
        o.rows[merge.s.r] = { cells: {} };
      }
      if (o.rows[merge.s.r].cells[merge.s.c] == null) {
        o.rows[merge.s.r].cells[merge.s.c] = {};
      }

      // 获取合并单元格的值
      const cellValue = ws[XLSX.utils.encode_cell(merge.s)]?.w;

      // 设置合并单元格的文本值
      o.rows[merge.s.r].cells[merge.s.c].text = cellValue || "";

      o.rows[merge.s.r].cells[merge.s.c].merge = [
        merge.e.r - merge.s.r,
        merge.e.c - merge.s.c,
      ];

      o.merges[i] = XLSX.utils.encode_range(merge);
    });
    o.sheetConfig = { gridLine: !gridStatus };
    out.push(o);
  });

  return out;
};

const rgbaToRgb = (hexColor) => {
  // Assuming a white background, so the background RGB is (255, 255, 255)
  const backgroundR = 255,
    backgroundG = 255,
    backgroundB = 255;

  // Extract RGBA from hex
  let r = parseInt(hexColor.slice(1, 3), 16);
  let g = parseInt(hexColor.slice(3, 5), 16);
  let b = parseInt(hexColor.slice(5, 7), 16);
  let a = parseInt(hexColor.slice(7, 9), 16) / 255.0; // Convert alpha to a scale of 0 to 1

  // Calculate new RGB by blending the original color with the background
  let newR = Math.round((1 - a) * backgroundR + a * r);
  let newG = Math.round((1 - a) * backgroundG + a * g);
  let newB = Math.round((1 - a) * backgroundB + a * b);

  // Convert RGB back to hex
  let newHexColor =
    "#" + ((1 << 24) + (newR << 16) + (newG << 8) + newB).toString(16).slice(1);

  return newHexColor.toUpperCase(); // Convert to uppercase as per original Python function
};

const getNewSheetName = (name, existingNames) => {
  let numericPart = name.match(/\d+$/);
  let baseName = name.replace(/\d+$/, "");

  if (!numericPart) {
    numericPart = "1";
  } else {
    numericPart = String(parseInt(numericPart[0], 10) + 1);
  }

  let newName = baseName + numericPart;

  while (existingNames.includes(newName)) {
    numericPart = String(parseInt(numericPart, 10) + 1);
    newName = baseName + numericPart;
  }

  return newName;
};

const getRowHeightForTextWrap = (ctx, textWrap, biw, text, fontSize) => {
  const txts = `${text}`.split("\n");
  const ntxts = [];
  txts.forEach((it) => {
    const txtWidth = ctx.measureText(it).width;
    if (textWrap && txtWidth > npx(biw)) {
      let textLine = { w: 0, len: 0, start: 0 };
      for (let i = 0; i < it.length; i += 1) {
        if (textLine.w >= npx(biw)) {
          ntxts.push(it.substr(textLine.start, textLine.len));
          textLine = { w: 0, len: 0, start: i };
        }
        textLine.len += 1;
        textLine.w += ctx.measureText(it[i]).width + 1;
      }
      if (textLine.len > 0) {
        ntxts.push(it.substr(textLine.start, textLine.len));
      }
    } else {
      ntxts.push(it);
    }
  });
  const rowHeight = ntxts.length * (fontSize + 2);
  return rowHeight;
};

const deepClone = (data) => JSON.parse(JSON.stringify(data));

const getNumberFormatFromStyles = (styleTag) => {
  const styleContent = styleTag.innerHTML;
  let match;
  const results = {};

  while (
    (match = EXTRACT_MSO_NUMBER_FORMAT_REGEX.exec(styleContent)) !== null
  ) {
    const className = match[1].replace("\n\t", "");
    const msoNumberFormat = match[2]
      .trim()
      .replaceAll("\\", "")
      .replaceAll("0022", "")
      .replaceAll('"', "");
    results[className] = msoNumberFormat?.slice(0, -1)?.trim() ?? "";
  }
  return results;
};

function generateSSFFormat(
  groupingSymbol = ",",
  digitGrouping = "",
  decimalUpto = 2,
  customFormat = "general"
) {
  let formatString = "";

  switch (customFormat.toLowerCase()) {
    case "text":
      formatString = "@"; // Text format
      break;
    case "number":
      formatString = generateNumberFormat(
        groupingSymbol,
        digitGrouping,
        decimalUpto
      );
      break;
    case "percent":
      formatString = generateNumberFormat(
        groupingSymbol,
        digitGrouping,
        decimalUpto
      );
      formatString = formatString.replace(/0(?=[);])/g, "0%"); // Ensures proper percentage formatting

      break;
    case "rmb":
      formatString =
        '"¥"' +
        generateNumberFormat(groupingSymbol, digitGrouping, decimalUpto);
      break;
    case "usd":
      formatString =
        '"$"' +
        generateNumberFormat(groupingSymbol, digitGrouping, decimalUpto);
      break;
    case "eur":
      formatString =
        '"€"' +
        generateNumberFormat(groupingSymbol, digitGrouping, decimalUpto);
      break;
    case "date":
      formatString = "yyyy-mm-dd";
      break;
    case "time":
      formatString = "hh:mm:ss";
      break;
    case "datetime":
      formatString = "yyyy-mm-dd hh:mm:ss";
      break;
    case "duration":
      formatString = "[hh]:mm:ss";
      break;
    default:
      formatString = generateNumberFormat(
        groupingSymbol,
        digitGrouping,
        decimalUpto
      );
      break;
  }

  return formatString;
}

function generateNumberFormat(groupingSymbol, digitGrouping, decimalUpto) {
  const integerPart = digitGrouping?.split(".")[0] ?? "";
  const groupingPositions = integerPart.split(groupingSymbol);
  const primaryGroupingSize =
    groupingPositions?.length > 1 ? groupingPositions[1].length : 3;

  let groupingPart = "#";
  if (primaryGroupingSize === 3) {
    groupingPart = "#,##0";
  } else {
    groupingPart = "#".repeat(primaryGroupingSize - 1) + "##0";
  }

  if (groupingSymbol !== ",") {
    groupingPart = groupingPart.replace(/,/g, groupingSymbol);
  }

  let decimalPart = "";
  if (decimalUpto > 0) {
    decimalPart = "." + "0".repeat(decimalUpto);
  }

  const formatString = groupingPart + decimalPart;

  // Add handling for negative values and zero
  const negativePart = `(${groupingPart + decimalPart})`; // Enclose negative numbers in parentheses
  const zeroPart = "-"; // Show zero as a dash
  // Return the final format string, including positive, negative, and zero formats
  return `${formatString};${negativePart};${zeroPart}`;
}

const countDecimals = (num) => {
  if (num.includes(".")) {
    return num.split(".")[1].length; // Count digits after decimal
  }
  return 0; // No decimal places
};

const getMultiplierPenalty = (numString) => {
  let num = Number(numString);
  if (num === 0) return 0; // Edge case for zero

  let count = 0;
  while (num % 10 === 0) {
    count++;
    num /= 10;
  }
  return count;
};

const getDecimalPlaces = (expression) => {
  let numbers = expression.match(/[\d.]+/g); // Extract numbers
  if (!numbers) return "Invalid Input";

  let totalDecimalPlaces = numbers.reduce(
    (sum, num) => sum + countDecimals(num),
    0
  );
  let totalPenalty = numbers.reduce(
    (sum, num) => sum + getMultiplierPenalty(num),
    0
  );

  return Math.max(totalDecimalPlaces - totalPenalty, 0);
};

export {
  getStylingForClass,
  parseCssToXDataStyles,
  parseHtmlToText,
  generateUniqueId,
  replaceCellRefWithNew,
  readExcelFile,
  stox,
  rgbaToRgb,
  getNewSheetName,
  getRowHeightForTextWrap,
  deepClone,
  getNumberFormatFromStyles,
  generateSSFFormat,
  getDecimalPlaces,
};
