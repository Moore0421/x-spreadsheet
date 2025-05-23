/* global document */

import Selector from "./selector";
import Scroll from "./scroll";
import History from "./history";
import Clipboard from "./clipboard";
import AutoFilter from "./auto_filter";
import { Merges } from "./merge";
import helper from "./helper";
import { Rows } from "./row";
import { Cols } from "./col";
import { Validations } from "./validation";
import { CellRange } from "./cell_range";
import { expr2expr, expr2xy, xy2expr } from "./alphabet";
import { t } from "../locale/locale";
import {
  generateUniqueId,
  getNumberFormatFromStyles,
  getRowHeightForTextWrap,
  getStylingForClass,
  parseCssToXDataStyles,
  parseHtmlToText,
  replaceCellRefWithNew,
} from "../utils";
import SheetConfig from "./sheetConfig";
import CellConfig from "./cellConfig";
import Variable from "./variable";
import {
  ACCOUNTING_FORMAT_REGEX,
  DEFAULT_ROW_HEIGHT,
  EXCEL_ERRORS,
  EXTRACT_FORMULA_CELL_NAME_REGEX,
} from "../constants";
import { getFontSizePxByPt } from "./font";
import { getDrawBox } from "../component/table";
import { npx } from "../canvas/draw";
import RowIdModal from "../component/toolbar/setting_row_id";
import { evaluateDynamicExpression } from "./cell";

// private methods
/*
 * {
 *  name: ''
 *  freeze: [0, 0],
 *  formats: [],
 *  styles: [
 *    {
 *      bgcolor: '',
 *      align: '',
 *      valign: '',
 *      textwrap: false,
 *      strike: false,
 *      underline: false,
 *      color: '',
 *      format: 1,
 *      border: {
 *        left: [style, color],
 *        right: [style, color],
 *        top: [style, color],
 *        bottom: [style, color],
 *      },
 *      font: {
 *        name: 'Helvetica',
 *        size: 10,
 *        bold: false,
 *        italic: false,
 *      }
 *    }
 *  ],
 *  merges: [
 *    'A1:F11',
 *    ...
 *  ],
 *  rows: {
 *    1: {
 *      height: 50,
 *      style: 1,
 *      cells: {
 *        1: {
 *          style: 2,
 *          type: 'string',
 *          text: '',
 *          value: '', // cal result
 *        }
 *      }
 *    },
 *    ...
 *  },
 *  cols: {
 *    2: { width: 100, style: 1 }
 *  }
 * }
 */
const defaultSettings = {
  mode: "edit", // edit | read
  view: {
    height: () => document.documentElement.clientHeight,
    width: () => document.documentElement.clientWidth,
  },
  showGrid: true,
  showToolbar: true,
  showContextmenu: true,
  showBottomBar: true,
  row: {
    len: 40,
    height: DEFAULT_ROW_HEIGHT,
    minHeight: DEFAULT_ROW_HEIGHT,
  },
  col: {
    len: 40,
    width: 100,
    indexWidth: 60,
    minWidth: 60,
  },
  style: {
    bgcolor: "#ffffff",
    align: "left",
    valign: "middle",
    textwrap: false,
    strike: false,
    underline: false,
    color: "#0a0a0a",
    font: {
      name: "Arial",
      size: 11,
      bold: false,
      italic: false,
    },
    format: "general",
  },
};

const toolbarHeight = 41;
const bottombarHeight = 41;

// src: cellRange
// dst: cellRange
function canPaste(src, dst, error = () => {}) {
  const { merges } = this;
  const cellRange = dst.clone();
  const [srn, scn] = src.size();
  const [drn, dcn] = dst.size();
  if (srn > drn) {
    cellRange.eri = dst.sri + srn - 1;
  }
  if (scn > dcn) {
    cellRange.eci = dst.sci + scn - 1;
  }
  if (merges.intersects(cellRange)) {
    error(t("error.pasteForMergedCell"));
    return false;
  }
  return true;
}
function copyPaste(srcCellRange, dstCellRange, what, autofill = false) {
  const { rows, merges } = this;
  // delete dest merge
  if (what === "all" || what === "format") {
    rows.deleteCells(dstCellRange, what);
    merges.deleteWithin(dstCellRange);
  }
  rows.copyPaste(srcCellRange, dstCellRange, what, autofill, (ri, ci, cell) => {
    if (cell && cell.merge) {

      const [rn, cn] = cell.merge;
      if (rn <= 0 && cn <= 0) return;
      merges.add(new CellRange(ri, ci, ri + rn, ci + cn));
    }
  });
}

function cutPaste(srcCellRange, dstCellRange) {
  const { clipboard, rows, merges } = this;
  rows.cutPaste(srcCellRange, dstCellRange);
  merges.move(
    srcCellRange,
    dstCellRange.sri - srcCellRange.sri,
    dstCellRange.sci - srcCellRange.sci
  );
  clipboard.clear();
}

// bss: { top, bottom, left, right }
function setStyleBorder(ri, ci, bss) {
  const { styles, rows } = this;
  const cell = rows.getCellOrNew(ri, ci);
  let cstyle = {};
  if (cell.style !== undefined) {
    cstyle = helper.cloneDeep(styles[cell.style]);
  }
  cstyle = helper.merge(cstyle, { border: bss });
  cell.style = this.addStyle(cstyle);
}

function setStyleBorders({ mode, style, color }) {
  const { styles, selector, rows } = this;
  const { sri, sci, eri, eci } = selector.range;
  const multiple = !this.isSingleSelected();
  if (!multiple) {
    if (mode === "inside" || mode === "horizontal" || mode === "vertical") {
      return;
    }
  }
  if (mode === "outside" && !multiple) {
    setStyleBorder.call(this, sri, sci, {
      top: [style, color],
      bottom: [style, color],
      left: [style, color],
      right: [style, color],
    });
  } else if (mode === "none") {
    selector.range.each((ri, ci) => {
      const cell = rows.getCell(ri, ci);
      if (cell && cell.style !== undefined) {
        const ns = helper.cloneDeep(styles[cell.style]);
        delete ns.border;
        // ['bottom', 'top', 'left', 'right'].forEach((prop) => {
        //   if (ns[prop]) delete ns[prop];
        // });
        cell.style = this.addStyle(ns);
      }
    });
  } else if (
    mode === "all" ||
    mode === "inside" ||
    mode === "outside" ||
    mode === "horizontal" ||
    mode === "vertical"
  ) {
    const merges = [];
    for (let ri = sri; ri <= eri; ri += 1) {
      for (let ci = sci; ci <= eci; ci += 1) {
        // jump merges -- start
        const mergeIndexes = [];
        for (let ii = 0; ii < merges.length; ii += 1) {
          const [mri, mci, rn, cn] = merges[ii];
          if (ri === mri + rn + 1) mergeIndexes.push(ii);
          if (mri <= ri && ri <= mri + rn) {
            if (ci === mci) {
              ci += cn + 1;
              break;
            }
          }
        }
        mergeIndexes.forEach((it) => merges.splice(it, 1));
        if (ci > eci) break;
        // jump merges -- end
        const cell = rows.getCell(ri, ci);
        let [rn, cn] = [0, 0];
        if (cell && cell.merge) {
          [rn, cn] = cell.merge;
          merges.push([ri, ci, rn, cn]);
        }
        const mrl = rn > 0 && ri + rn === eri;
        const mcl = cn > 0 && ci + cn === eci;
        let bss = {};
        if (mode === "all") {
          bss = {
            bottom: [style, color],
            top: [style, color],
            left: [style, color],
            right: [style, color],
          };
        } else if (mode === "inside") {
          if (!mcl && ci < eci) bss.right = [style, color];
          if (!mrl && ri < eri) bss.bottom = [style, color];
        } else if (mode === "horizontal") {
          if (!mrl && ri < eri) bss.bottom = [style, color];
        } else if (mode === "vertical") {
          if (!mcl && ci < eci) bss.right = [style, color];
        } else if (mode === "outside" && multiple) {
          if (sri === ri) bss.top = [style, color];
          if (mrl || eri === ri) bss.bottom = [style, color];
          if (sci === ci) bss.left = [style, color];
          if (mcl || eci === ci) bss.right = [style, color];
        }
        if (Object.keys(bss).length > 0) {
          setStyleBorder.call(this, ri, ci, bss);
        }
        ci += cn;
      }
    }
  } else if (mode === "top" || mode === "bottom") {
    for (let ci = sci; ci <= eci; ci += 1) {
      if (mode === "top") {
        setStyleBorder.call(this, sri, ci, { top: [style, color] });
        ci += rows.getCellMerge(sri, ci)[1];
      }
      if (mode === "bottom") {
        setStyleBorder.call(this, eri, ci, { bottom: [style, color] });
        ci += rows.getCellMerge(eri, ci)[1];
      }
    }
  } else if (mode === "left" || mode === "right") {
    for (let ri = sri; ri <= eri; ri += 1) {
      if (mode === "left") {
        setStyleBorder.call(this, ri, sci, { left: [style, color] });
        ri += rows.getCellMerge(ri, sci)[0];
      }
      if (mode === "right") {
        setStyleBorder.call(this, ri, eci, { right: [style, color] });
        ri += rows.getCellMerge(ri, eci)[0];
      }
    }
  }
}

function getCellRowByY(y, scrollOffsety) {
  const { rows } = this;
  const fsh = this.freezeTotalHeight();

  let inits = rows.height;
  if (fsh + rows.height < y) inits -= scrollOffsety;

  // handle ri in autofilter
  const frset = this.exceptRowSet;

  let ri = 0;
  let top = inits;
  let { height } = rows;
  for (; ri < rows.len; ri += 1) {
    if (top > y) break;
    if (!frset.has(ri)) {
      height = rows.getHeight(ri);
      top += height;
    }
  }
  top -= height;


  if (top <= 0) {
    return { ri: -1, top: 0, height };
  }

  return { ri: ri - 1, top, height };
}

function getCellColByX(x, scrollOffsetx) {
  const { cols } = this;
  const fsw = this.freezeTotalWidth();
  let inits = cols.indexWidth;
  if (fsw + cols.indexWidth < x) inits -= scrollOffsetx;
  const [ci, left, width] = helper.rangeReduceIf(
    0,
    cols.len,
    inits,
    cols.indexWidth,
    x,
    (i) => cols.getWidth(i)
  );
  if (left <= 0) {
    return { ci: -1, left: 0, width: cols.indexWidth };
  }
  return { ci: ci - 1, left, width };
}

export default class DataProxy {
  constructor(parent, name, settings) {
    this.parent = parent;
    this.name = name;
    this.settings = helper.merge(defaultSettings, settings || {});
    // 初始化时如果没有mode，设置为normal
    if (!this.settings.mode) {
      this.settings.mode = 'normal';
    }
    // save data begin
    this.freeze = [0, 0];
    this.styles = []; // Array<Style>
    this.merges = new Merges(); // [CellRange, ...]
    this.rows = new Rows(this, this.settings.row, this.settings);
    this.cols = new Cols(this.settings.col);
    this.sheetConfig = new SheetConfig(this.name, this.settings);
    this.cellConfig = new CellConfig(this.name, this.settings);
    this.validations = new Validations();
    this.hyperlinks = {};
    this.comments = {};
    this.variables = new Variable(parent);
    // save data end

    // don't save object
    this.selector = new Selector();
    this.scroll = new Scroll();
    this.history = new History();
    this.clipboard = new Clipboard();
    this.autoFilter = new AutoFilter();
    this.change = () => {};
    this.includeRowSet = new Set();
    this.exceptRowSet = new Set();
    this.sortedRowMap = new Map();
    this.unsortedRowMap = new Map();
    this.rootContext = parent;
    this.sheetId = this.sheetConfig.sheetId ?? generateUniqueId();
  }

  addValidation(mode, ref, validator) {

    this.changeData(() => {
      this.validations.add(mode, ref, validator);
    });
  }

  removeValidation() {
    const { range } = this.selector;
    this.changeData(() => {
      this.validations.remove(range);
    });
  }

  getSelectedValidator() {
    const { ri, ci } = this.selector;
    const v = this.validations.get(ri, ci);
    return v ? v.validator : null;
  }

  getSelectedValidation() {
    const { ri, ci, range } = this.selector;
    const v = this.validations.get(ri, ci);
    const ret = { ref: range.toString() };
    if (v !== null) {
      ret.mode = v.mode;
      ret.validator = v.validator;
    }
    return ret;
  }

  canUndo() {
    return this.history.canUndo();
  }

  canRedo() {
    return this.history.canRedo();
  }

  undo() {
    this.history.undo(this.getData(), (d) => {
      this.setData(d);
    });
  }

  redo() {
    this.history.redo(this.getData(), (d) => {
      this.setData(d);
    });
  }

  copy() {
    this.clipboard.copy(this.selector.range);
  }

  copyToSystemClipboard(evt) {
    let copyText = [];
    const { sri, eri, sci, eci } = this.selector.range;

    for (let ri = sri; ri <= eri; ri += 1) {
      const row = [];
      for (let ci = sci; ci <= eci; ci += 1) {
        const cell = this.getCell(ri, ci);
        row.push((cell && cell.text) || "");
      }
      copyText.push(row);
    }

    // Adding \n and why not adding \r\n is to support online office and client MS office and WPS
    copyText = copyText.map((row) => row.join("\t")).join("\n");

    // why used this
    // cuz http protocol will be blocked request clipboard by browser
    if (evt) {
      evt.clipboardData.clearData();
      evt.clipboardData.setData("text/plain", copyText);
      evt.preventDefault();
    }

    // this need https protocol
    /* global navigator */
    if (navigator.clipboard) {
      navigator.clipboard.writeText(copyText).then(
        () => {},
        (err) => {
          console.error("Failed to write clipboard contents: ", err);
        }
      );
    }
  }

  cut() {
    this.clipboard.cut(this.selector.range);
  }

  // what: all | text | format
  paste(what = "all", error = () => {}) {

    const { clipboard, selector } = this;
    if (clipboard.isClear()) return false;
    if (!canPaste.call(this, clipboard.range, selector.range, error))
      return false;

    this.changeData(() => {
      if (clipboard.isCopy()) {
        copyPaste.call(this, clipboard.range, selector.range, what);
      } else if (clipboard.isCut()) {
        cutPaste.call(this, clipboard.range, selector.range);
      }
    });
    return true;
  }

  pasteFromSystemClipboard(evt, resetSheet) {
    try {
      const { selector } = this;
      const htmlContent = evt?.clipboardData?.getData("text/html");
      navigator.clipboard
        .read()
        .then((data) => {
          data.forEach((item) => {
            if (item?.types?.includes("text/html")) {
              const parser = new DOMParser();
              const htmlDocument = parser.parseFromString(
                htmlContent,
                "text/html"
              );
              const htmlStyles =
                htmlDocument?.getElementsByTagName("style")?.[0];
              const htmlBody = htmlDocument?.getElementsByTagName("body")?.[0];
              const tableContent = htmlBody?.querySelectorAll("table");
              const numberFormats = getNumberFormatFromStyles(htmlStyles);

              tableContent?.forEach((table) => {
                const tBody = table?.querySelector("tbody");
                const trs = tBody?.querySelectorAll("tr");
                let startRow = selector.ri;
                trs?.forEach((tr) => {
                  let startColumn = selector.ci;
                  const tds = tr?.querySelectorAll("td");
                  tds?.forEach((td) => {
                    const cellContent = parseHtmlToText(td?.innerHTML ?? "");
                    let rawCellContent = cellContent;
                    let isNegative = false;
                    let isAccountingFormat = false;
                    if (cellContent[0] === "-") {
                      isNegative = true;
                      rawCellContent = cellContent.substring(
                        1,
                        cellContent.length
                      );
                    }
                    if (ACCOUNTING_FORMAT_REGEX.test(rawCellContent)) {
                      isAccountingFormat = true;
                      isNegative =
                        isNegative ||
                        rawCellContent.includes("(") ||
                        rawCellContent.includes("-");
                      rawCellContent = rawCellContent.replace(/[^\d.]/g, "");
                      if (isNegative) rawCellContent = `-${rawCellContent}`;
                    }

                    const mergedCellRange = this.merges.getFirstIncludes(
                      startRow,
                      startColumn
                    );

                    if (mergedCellRange) {
                      for (
                        let k = startColumn;
                        k <= mergedCellRange.eci;
                        k += 1
                      )
                        startColumn += 1;
                    }

                    this.setCellText(
                      startRow,
                      startColumn,
                      isAccountingFormat ? rawCellContent : cellContent,
                      "input"
                    );
                    const rowSpan = Number(td?.getAttribute("rowspan") ?? 1);
                    const colSpan = Number(td?.getAttribute("colspan") ?? 1);
                    if (rowSpan - 1 !== 0 || colSpan - 1 !== 0) {
                      this.setCellMerge(startRow, startColumn, [
                        rowSpan - 1,
                        colSpan - 1,
                      ]);
                    }
                    const cellClassName = td?.getAttribute("class");
                    const cellStyleString = `${td?.getAttribute("style") ?? ""};${getStylingForClass(htmlStyles, cellClassName)}`;
                    const cellStyle =
                      parseCssToXDataStyles(cellStyleString)?.parsedStyles ??
                      {};
                    const { width, height } = cellStyle?.dimensions ?? {};
                    delete cellStyle.dimensions;
                    if (width && this.getColWidth(startColumn) < width) {
                      this.setColWidth(startColumn, width);
                    }
                    if (height && this.getRowHeight(startRow) < height) {
                      this.setRowHeight(startRow, height);
                    }
                    const cell = this.getCell(startRow, startColumn);
                    if (cell) {
                      if (numberFormats.hasOwnProperty(cellClassName)) {
                        cell.t = "n";
                        cell.z = numberFormats[cellClassName];
                        cell.w = cellContent;
                        cellStyle.align = "right";
                        if (cellContent === "-") cell.text = 0;
                      } else if (!isNaN(Number(rawCellContent))) {
                        cell.t = "n";
                        cell.w = cellContent;
                        cellStyle.align = "right";
                        if (cellContent === "-") cell.text = 0;
                      }
                      cell.style = this.addStyle(cellStyle);
                    }

                    startColumn += colSpan;
                  });
                  startRow += 1;
                });
              });
              resetSheet();
            } else if (item?.types?.includes("text/plain")) {
              navigator.clipboard.readText().then((content) => {
                const contentToPaste = this.parseClipboardContent(content);
                let startRow = selector.ri;
                contentToPaste.forEach((row) => {
                  let startColumn = selector.ci;
                  row.forEach((cellContent) => {
                    this.setCellText(
                      startRow,
                      startColumn,
                      cellContent,
                      "input"
                    );
                    startColumn += 1;
                  });
                  startRow += 1;
                });
                resetSheet();
              });
            }
          });
        })
        .catch((err) => {
          console.error("Failed to read clipboard contents: ", err);
        });
    } catch (err) {
      console.error("Pasting error: ", err);
    }
  }

  parseClipboardContent(clipboardContent) {
    const parsedData = [];

    // first we need to figure out how many rows we need to paste
    const rows = clipboardContent.split("\n");

    // for each row parse cell data
    let i = 0;
    rows.forEach((row) => {
      parsedData[i] = row.split("\t");
      i += 1;
    });
    return parsedData;
  }

  pasteFromText(txt) {
    let lines = [];

    if (/\r\n/.test(txt))
      lines = txt.split("\r\n").map((it) => it.replace(/"/g, "").split("\t"));
    else lines = txt.split("\n").map((it) => it.replace(/"/g, "").split("\t"));

    if (lines.length) {
      const { rows, selector } = this;

      this.changeData(() => {
        rows.paste(lines, selector.range);
      });
    }
  }

  autofill(cellRange, what, error = () => {}) {
    const srcRange = this.selector.range;
    if (!canPaste.call(this, srcRange, cellRange, error)) return false;
    this.changeData(() => {
      copyPaste.call(this, srcRange, cellRange, what, true);
    });
    return true;
  }

  clearClipboard() {
    this.clipboard.clear();
  }

  calSelectedRangeByEnd(ri, ci) {
    const { selector, rows, cols, merges } = this;
    let { sri, sci, eri, eci } = selector.range;
    const cri = selector.ri;
    const cci = selector.ci;
    let [nri, nci] = [ri, ci];
    if (ri < 0) nri = rows.len - 1;
    if (ci < 0) nci = cols.len - 1;
    if (nri > cri) [sri, eri] = [cri, nri];
    else [sri, eri] = [nri, cri];
    if (nci > cci) [sci, eci] = [cci, nci];
    else [sci, eci] = [nci, cci];
    if (ri !== -1 && ci !== -1) {
      selector.range = merges.union(new CellRange(sri, sci, eri, eci));
      selector.range = merges.union(selector.range);
    }
    return selector.range;
  }

  getMaximumAvailableIndexes() {
    const { rows } = this;
    let maxRow = 0,
      maxCol = 0;
    const rowList = rows?._;
    if (rowList) {
      for (const rowKey of Object.keys(rowList)) {
        const numericRowKey = Number(rowKey);
        const cells = rowList[rowKey].cells ?? {};
        if (maxRow <= numericRowKey && Object.keys(cells)?.length) {
          maxRow = numericRowKey;
        }

        for (const cellKey of Object.keys(cells)) {
          const numericCellKey = Number(cellKey);
          if (maxCol <= numericCellKey) {
            maxCol = numericCellKey;
          }
        }
      }
    }
    return { maxRow, maxCol };
  }

  calSelectedRangeByStart(ri, ci) {
    const { selector, rows, cols, merges, settings = {} } = this;
    const { suppressMaximumSelection } = settings;
    let maxRow = 0,
      maxCol = 0;
    if (suppressMaximumSelection) {
      ({ maxRow, maxCol } = this.getMaximumAvailableIndexes());
    }
    let cellRange = merges.getFirstIncludes(ri, ci);
    if (cellRange === null) {
      cellRange = new CellRange(ri, ci, ri, ci);
      if (ri === -1) {
        cellRange.sri = 0;
        cellRange.eri = suppressMaximumSelection ? maxRow : rows.len - 1;
      }
      if (ci === -1) {
        cellRange.sci = 0;
        cellRange.eci = suppressMaximumSelection ? maxCol : cols.len - 1;
      }
    }
    selector.range = cellRange;
    return cellRange;
  }

  setSelectedCellAttr(property, value) {
    this.changeData(() => {
      const { selector, styles, rows, sheetConfig, cellConfig } = this;
      if (property === "merge") {
        if (value) this.merge();
        else this.unmerge();
      } else if (property === "border") {
        setStyleBorders.call(this, value);
      } else if (property === "formula") {

        const { ri, ci, range } = selector;
        if (selector.multiple()) {
          const [rn, cn] = selector.size();
          const { sri, sci, eri, eci } = range;
          if (rn > 1) {
            for (let i = sci; i <= eci; i += 1) {
              const cell = rows.getCellOrNew(eri + 1, i);
              cell.text = `=${value}(${xy2expr(i, sri)}:${xy2expr(i, eri)})`;
            }
          } else if (cn > 1) {
            const cell = rows.getCellOrNew(ri, eci + 1);
            cell.text = `=${value}(${xy2expr(sci, ri)}:${xy2expr(eci, ri)})`;
          }
        } else {
          const cell = rows.getCellOrNew(ri, ci);
          cell.text = `=${value}()`;
        }
      } else if (property === "grid") {
        sheetConfig.setData({ gridLine: value });
      } else if (
        cellConfig?.cellButtons?.find((button) => button?.tag === property)
      ) {
        selector.range.each((ri, ci) => {
          const cell = rows.getCellOrNew(ri, ci);
          cell.cellMeta = cell.cellMeta ?? {};
          cell.cellMeta[property] = value;
        });
      } else {
        selector.range.each((ri, ci) => {
          const cell = rows.getCellOrNew(ri, ci);
          let cstyle = {};
          if (cell.style !== undefined) {
            cstyle = helper.cloneDeep(styles[cell.style]);
          }
          if (property === "format") {
            cstyle.format = value;
            cell.style = this.addStyle(cstyle);
          } else if (
            property === "font-bold" ||
            property === "font-italic" ||
            property === "font-name" ||
            property === "font-size"
          ) {
            const nfont = {};
            nfont[property.split("-")[1]] = value;
            cstyle.font = Object.assign(cstyle.font || {}, nfont);
            cell.style = this.addStyle(cstyle);
          } else if (
            property === "strike" ||
            property === "textwrap" ||
            property === "underline" ||
            property === "align" ||
            property === "valign" ||
            property === "color" ||
            property === "bgcolor"
          ) {
            if (property === "textwrap") {
              const cellStyle = this.getCellStyleOrDefault(ri, ci);
              const drawBox = getDrawBox(this, ri, ci);
              const draw = this.rootContext.sheet.table.draw;
              const font = Object.assign({}, cellStyle.font);
              font.size = getFontSizePxByPt(font.size);
              draw?.attr({
                align: cellStyle.align,
                valign: cellStyle.valign,
                font: `${font.italic ? "italic" : ""} ${font.bold ? "bold" : ""} ${npx(font.size)}px ${font.name}`,
                color: cellStyle.color,
                strike: cellStyle.strike,
                underline: cellStyle.underline,
              });

              const oldRowHeight = this.rows.getHeight(ri);
              const newRowHeight = value
                ? getRowHeightForTextWrap(
                    draw.ctx,
                    value,
                    drawBox.innerWidth(),
                    cell.text,
                    font.size
                  )
                : DEFAULT_ROW_HEIGHT;

              const rowHeight = value
                ? oldRowHeight > newRowHeight
                  ? oldRowHeight
                  : newRowHeight > DEFAULT_ROW_HEIGHT
                    ? newRowHeight
                    : DEFAULT_ROW_HEIGHT
                : DEFAULT_ROW_HEIGHT;
              if (!this.rows.isHeightChanged)
                this.rows.setHeight(ri, rowHeight, true);
            }
            cstyle[property] = value;
            cell.style = this.addStyle(cstyle);
          } else if (property === "editable" || this.settings.mode === 'design') {
            cell[property] = value;
            cell.cellMeta = cell?.cellMeta ?? {};
            cell.cellMeta[property] = value;
          } else {
            cell[property] = value;
          }
        });
      }
    });
  }

  setCellState(ri, ci, editable = true) {
    const { rows } = this;
    const cell = rows.getCellOrNew(ri, ci);
    cell.editable = editable;
  }
  extractCellReferences(formula) {
    const cellRefRegex = /([A-Z]+[1-9][0-9]*)/g; // Regular expression to match cell references (e.g., A1, B10)
    const matches = formula.match(cellRefRegex) || [];
    return matches;
  }
  // state: input | finished
  setFormulaCellText(text, ri, ci, state = "input") {

    const { autoFilter, rows } = this;

    // 定义运算符列表
    const operators = ['+', '-', '*', '/', '===', '==', '!==', '!=', '>=', '<=', '>', '<', '&&', '||', '%', '?', 'return'];
    
    // 优先处理$$前缀，确保不会被误判为其他类型的表达式
    const isJavaScript = text && text.startsWith('$$');
    
    // 只有非JavaScript代码才考虑是否为JS运算表达式
    const isJSExpression = !isJavaScript && text && text.startsWith('$') && 
      (operators.some(op => text.includes(op)) || 
       // 检查是否包含常见的字符串连接模式
       /\s*\+\s*['"]/.test(text));
    
    // 只有既不是JavaScript代码也不是JS运算表达式的$开头表达式才是动态函数
    const isDynamicExpression = !isJavaScript && !isJSExpression && 
      text && text.startsWith('$');

    if (state === "finished") {
      if (isJavaScript) {
        // 保存原始$$表达式到cell.formula
        const cell = rows.getCellOrNew(ri, ci);
        cell.formula = text;
        cell.isJavaScript = true;
        // 设置属性以支持双击编辑
        rows.setCellProperty(ri, ci, "f", text);
      } else if (isJSExpression) {
        // 处理JS运算表达式

        
        // 保存原始表达式
        const cell = rows.getCellOrNew(ri, ci);
        cell.formula = text;
        cell.isJSExpression = true;
        
        // 计算表达式结果
        try {
          // 去掉开头的$符号再计算
          const jsExpressionResult = evaluateJsExpression(text.substring(1), this);

          
          // 设置单元格显示值
          cell.text = jsExpressionResult?.toString() || '';
          
          // 设置属性以支持双击编辑
          rows.setCellProperty(ri, ci, "f", text);
        } catch (error) {
          console.error("JS运算表达式计算错误:", error);
          cell.text = `#JS_EXPRESSION_ERROR! ${error.message}`;
          if (!cell.style) cell.style = {};
          cell.style.color = "#FF0000"; // 错误标红
        }
      } else if (isDynamicExpression) {
        // 动态函数表达式处理逻辑

        
        // 保存原始$表达式
        const cell = rows.getCellOrNew(ri, ci);
        // 将原始表达式存储到 cell.formula
        cell.formula = text;
        // 记录单元格使用了动态表达式，便于双击时编辑原始表达式
        cell.isDynamicExpression = true;

        // 计算值并设置错误样式
        const value = evaluateDynamicExpression(text.substring(1), this);

        
        // 处理结果
        if (value && typeof value === "string" && value.startsWith("#") && value.endsWith("!")) {
          // 这是一个错误值，应用错误样式
          if (!cell.style) cell.style = {};
          cell.style.color = "#FF0000"; // 红色错误文本
          console.error("动态函数执行错误:", value);
          cell.text = value;
        } else {
          // 正常值，确保没有错误样式
          if (cell.style && cell.style.color === "#FF0000") {
            delete cell.style.color;
            if (Object.keys(cell.style).length === 0) {
              delete cell.style;
            }
          }

          // 将计算结果赋值给 cell.text
          cell.text = value?.toString() || '';
        }

        // 特殊处理动态函数表达式的updatedFormula
        // 只将真正的单元格引用转为大写，保留其他标识符的大小写
        let updatedFormula = text;
        
        // 仅转换独立的单元格引用，如A1、B2等
        // 创建一个只匹配独立单元格引用的正则表达式
        const cellRefRegex = /\b([A-Z]+[0-9]+)\b(?!\()/gi;
        updatedFormula = updatedFormula.replace(cellRefRegex, match => match.toUpperCase());
        

        rows.setCellProperty(ri, ci, "f", updatedFormula);
      } else if (text && text[0] === "=") {
        // 现有公式处理逻辑
        const isFormula = text[0] === "=";
        const updatedFormula =
          text?.replace(EXTRACT_FORMULA_CELL_NAME_REGEX, (match) =>
            match.toUpperCase()
          ) ?? text;

        const { onFormulaCellFinalized } = this.settings;
      if (onFormulaCellFinalized) {
        const cell = rows.getCellOrNew(ri, ci);
        const cellRefs = this.extractCellReferences(updatedFormula);
        onFormulaCellFinalized({ ...this, updatedFormula, cell, cellRefs });
      }

      rows.setCellProperty(ri, ci, "f", isFormula ? updatedFormula : text);
      } else {
        // 普通文本处理逻辑 - 直接设置文本值
        rows.setCellText(ri, ci, text);
        
        // 确保清除任何可能存在的动态表达式标记
        const cell = rows.getCell(ri, ci);
        if (cell) {
          delete cell.isDynamicExpression;
          delete cell.isJSExpression;
          delete cell.isJavaScript;
          delete cell.formula;
        }
      }
    } else {
      // 输入状态处理逻辑

      let nri = ri;
      if (this.unsortedRowMap.has(ri)) {
        nri = this.unsortedRowMap.get(ri);
      }
      const oldCell = rows.getCell(nri, ci);
      const oldText = oldCell ? oldCell.text : "";
      this.setCellText(nri, ci, text, state);
      // replace filter.value
      if (autoFilter.active()) {
        const filter = autoFilter.getFilter(ci);
        if (filter) {
          const vIndex = filter.value.findIndex((v) => v === oldText);
          if (vIndex >= 0) {
            filter.value.splice(vIndex, 1, text);
          }
        }
      }
    }
  }

  // 获取单元格编辑文本，需要返回原始表达式
  getCellEditText(ri, ci) {
    const { rows } = this;
    const cell = rows.getCell(ri, ci);

    // 如果是动态表达式，返回原始$表达式
    if (cell?.isDynamicExpression) {
      return cell.formula;
    }

    // 现有逻辑
    return cell?.text || "";
  }

  // 获取单元格编辑文本，需要返回原始表达式
  getCellEditText(ri, ci) {
    const { rows } = this;
    const cell = rows.getCell(ri, ci);

    // 如果是动态表达式，返回原始$表达式
    if (cell?.isDynamicExpression) {
      return cell.formula;
    }

    // 现有逻辑
    return cell?.text || "";
  }

  // state: input | finished
  setSelectedCellText(text, state = "input") {
    const { autoFilter, selector, rows } = this;
    const { ri, ci } = selector;
    let nri = ri;
    if (this.unsortedRowMap.has(ri)) {
      nri = this.unsortedRowMap.get(ri);
    }
    const oldCell = rows.getCell(nri, ci);
    const oldText = oldCell ? oldCell.text : "";
    this.setCellText(nri, ci, text, state);
    // replace filter.value
    if (autoFilter.active()) {
      const filter = autoFilter.getFilter(ci);
      if (filter) {
        const vIndex = filter.value.findIndex((v) => v === oldText);
        if (vIndex >= 0) {
          filter.value.splice(vIndex, 1, text);
        }
      }
    }
  }

  getSelectedCell() {
    const { ri, ci } = this.selector;
    let nri = ri;
    if (this.unsortedRowMap.has(ri)) {
      nri = this.unsortedRowMap.get(ri);
    }
    return this.rows.getCell(nri, ci);
  }

  xyInSelectedRect(x, y) {
    const { left, top, width, height } = this.getSelectedRect();
    const x1 = x - this.cols.indexWidth;
    const y1 = y - this.rows.height;

    return x1 > left && x1 < left + width && y1 > top && y1 < top + height;
  }

  getSelectedRect() {
    return this.getRect(this.selector.range);
  }

  getClipboardRect() {
    const { clipboard } = this;
    if (!clipboard.isClear()) {
      return this.getRect(clipboard.range);
    }
    return { left: -100, top: -100 };
  }

  getRect(cellRange) {
    const { scroll, rows, cols, exceptRowSet } = this;
    const { sri, sci, eri, eci } = cellRange;

    // no selector
    if (sri < 0 && sci < 0) {
      return {
        left: 0,
        l: 0,
        top: 0,
        t: 0,
        scroll,
      };
    }
    const left = cols.sumWidth(0, sci);
    const top = rows.sumHeight(0, sri, exceptRowSet);
    const height = rows.sumHeight(sri, eri + 1, exceptRowSet);
    const width = cols.sumWidth(sci, eci + 1);

    let left0 = left - scroll.x;
    let top0 = top - scroll.y;
    const fsh = this.freezeTotalHeight();
    const fsw = this.freezeTotalWidth();
    if (fsw > 0 && fsw > left) {
      left0 = left;
    }
    if (fsh > 0 && fsh > top) {
      top0 = top;
    }
    return {
      l: left,
      t: top,
      left: left0,
      top: top0,
      height,
      width,
      scroll,
    };
  }

  getCellRectByXY(x, y) {
    const { scroll, merges, rows, cols } = this;
    let { ri, top, height } = getCellRowByY.call(this, y, scroll.y);
    let { ci, left, width } = getCellColByX.call(this, x, scroll.x);
    if (ci === -1) {
      width = cols.totalWidth();
    }
    if (ri === -1) {
      height = rows.totalHeight();
    }
    if (ri >= 0 || ci >= 0) {
      const merge = merges.getFirstIncludes(ri, ci);
      if (merge) {
        ri = merge.sri;
        ci = merge.sci;
        ({ left, top, width, height } = this.cellRect(ri, ci));
      }
    }
    return {
      ri,
      ci,
      left,
      top,
      width,
      height,
    };
  }

  isSingleSelected() {
    const { sri, sci, eri, eci } = this.selector.range;
    const cell = this.getCell(sri, sci);
    if (cell && cell.merge) {
      const [rn, cn] = cell.merge;
      if (sri + rn === eri && sci + cn === eci) return true;
    }
    return !this.selector.multiple();
  }

  canUnmerge() {
    const { sri, sci, eri, eci } = this.selector.range;
    const cell = this.getCell(sri, sci);
    if (cell && cell.merge) {
      const [rn, cn] = cell.merge;
      if (sri + rn === eri && sci + cn === eci) return true;
    }
    return false;
  }

  merge() {
    const { selector, rows } = this;
    if (this.isSingleSelected()) return;
    const [rn, cn] = selector.size();

    if (rn > 1 || cn > 1) {
      const { sri, sci } = selector.range;
      this.changeData(() => {
        const cell = rows.getCellOrNew(sri, sci);
        cell.merge = [rn - 1, cn - 1];
        this.merges.add(selector.range);
        // delete merge cells
        this.rows.deleteCells(selector.range);

        this.rows.setCell(sri, sci, cell);
      });
    }
  }

  unmerge() {
    const { selector } = this;
    if (!this.isSingleSelected()) return;
    const { sri, sci } = selector.range;
    this.changeData(() => {
      this.rows.deleteCell(sri, sci, "merge");
      this.merges.deleteWithin(selector.range);
    });
  }

  canAutofilter() {
    return !this.autoFilter.active();
  }

  autofilter() {
    const { autoFilter, selector } = this;
    this.changeData(() => {
      if (autoFilter.active()) {
        autoFilter.clear();
        this.includeRowSet = new Set();
        this.exceptRowSet = new Set();
        this.sortedRowMap = new Map();
        this.unsortedRowMap = new Map();
      } else {
        autoFilter.ref = selector.range.toString();
      }
    });
  }

  setAutoFilter(ci, order, operator, value) {
    const { autoFilter } = this;
    autoFilter.addFilter(ci, operator, value);
    autoFilter.setSort(ci, order);
    this.resetAutoFilter();
  }

  resetAutoFilter() {
    const { autoFilter, rows } = this;
    if (!autoFilter.active()) return;
    const { sort } = autoFilter;
    const { rset, fset } = autoFilter.filteredRows((r, c) =>
      rows.getCell(r, c)
    );
    const fary = Array.from(fset);
    const oldAry = Array.from(fset);
    if (sort) {
      fary.sort((a, b) => {
        if (sort.order === "asc") return a - b;
        if (sort.order === "desc") return b - a;
        return 0;
      });
    }
    this.includeRowSet = fset;
    this.exceptRowSet = rset;
    this.sortedRowMap = new Map();
    this.unsortedRowMap = new Map();
    fary.forEach((it, index) => {
      this.sortedRowMap.set(oldAry[index], it);
      this.unsortedRowMap.set(it, oldAry[index]);
    });
  }

  deleteCell(what = "all") {
    const { selector } = this;
    this.changeData(() => {
      this.rows.deleteCells(selector.range, what);
      if (what === "all" || what === "format") {
        this.merges.deleteWithin(selector.range);
      }
    });
  }

  // type: row | column
  insert(type) {
    this.changeData(() => {
      const { sri, eri, sci, eci } = this.selector.range;
      const { rows, merges, cols } = this;
      const trigger = this?.rootContext?.sheet?.trigger;
      let si = sri;
      let n = 1;
      if (type === "row") {
        n = eri - sri + 1;
        rows.insert(sri, n);
        this.updateOtherSheetsFormulaOnInsert(type, sri, n);
        trigger?.call(this?.rootContext?.sheet, "row-column-operation", {
          on: type,
          type: "INSERT",
          count: n,
          startIndex: sri,
        });
      } else if (type === "column") {
        n = eci - sci + 1;
        rows.insertColumn(sci, n);
        this.updateOtherSheetsFormulaOnInsert(type, sci, n);
        si = sci;
        cols.len += n;
        Object.keys(cols._)
          .reverse()
          .forEach((colIndex) => {
            const col = parseInt(colIndex, 10);
            if (col >= sci) {
              cols._[col + n] = cols._[col];
              delete cols._[col];
            }
          });
        trigger?.call(this?.rootContext?.sheet, "row-column-operation", {
          on: type,
          type: "INSERT",
          count: n,
          startIndex: sci,
        });
      }
      merges.shift(type, si, n, (ri, ci, rn, cn) => {
        const cell = rows.getCell(ri, ci);
        cell.merge[0] += rn;
        cell.merge[1] += cn;
      });
    });
  }

  // type: row | column
  delete(type) {
    this.changeData(() => {
      const { rows, merges, selector, cols } = this;
      const { range } = selector;
      const { sri, sci, eri, eci } = selector.range;
      const trigger = this?.rootContext?.sheet?.trigger;
      const [rsize, csize] = selector.range.size();
      let si = sri;
      let size = rsize;
      if (type === "row") {
        rows.delete(sri, eri);
        this.updateOtherSheetsFormulaOnDelete(type, sri, eri);
        trigger?.call(this?.rootContext?.sheet, "row-column-operation", {
          on: type,
          type: "DELETE",
          count: eri - sri + 1,
          startIndex: sri,
        });
      } else if (type === "column") {
        rows.deleteColumn(sci, eci);
        this.updateOtherSheetsFormulaOnDelete(type, sci, eci);
        si = range.sci;
        size = csize;
        cols.len -= eci - sci + 1;
        Object.keys(cols._).forEach((colIndex) => {
          const col = parseInt(colIndex, 10);
          if (col >= sci) {
            if (col > eci) cols._[col - (eci - sci + 1)] = cols._[col];
            delete cols._[col];
          }
        });
        trigger?.call(this?.rootContext?.sheet, "row-column-operation", {
          on: type,
          type: "DELETE",
          count: eci - sci + 1,
          startIndex: sci,
        });
      }

      merges.shift(type, si, -size, (ri, ci, rn, cn) => {

        const cell = rows.getCell(ri, ci);
        cell.merge[0] += rn;
        cell.merge[1] += cn;
        if (cell.merge[0] === 0 && cell.merge[1] === 0) {
          delete cell.merge;
        }
      });
    });
  }

  scrollx(x, cb) {
    const { scroll, freeze, cols } = this;
    const [, fci] = freeze;
    const [ci, left, width] = helper.rangeReduceIf(
      fci,
      cols.len,
      0,
      0,
      x,
      (i) => cols.getWidth(i)
    );

    let x1 = left;
    if (x > 0) x1 += width;
    if (scroll.x !== x1) {
      scroll.ci = x > 0 ? ci : 0;
      scroll.x = x1;
      cb();
    }
  }

  scrolly(y, cb) {
    const { scroll, freeze, rows } = this;
    const [fri] = freeze;
    const [ri, top, height] = helper.rangeReduceIf(
      fri,
      rows.len,
      0,
      0,
      y,
      (i) => rows.getHeight(i)
    );
    let y1 = top;
    if (y > 0) y1 += height;

    if (scroll.y !== y1) {
      scroll.ri = y > 0 ? ri : 0;
      scroll.y = y1;
      cb();
    }
  }

  cellRect(ri, ci) {
    const { rows, cols } = this;
    const left = cols.sumWidth(0, ci);
    const top = rows.sumHeight(0, ri);
    const cell = rows.getCell(ri, ci);
    let width = cols.getWidth(ci);
    let height = rows.getHeight(ri);
    if (cell !== null) {
      if (cell.merge) {
        const [rn, cn] = cell.merge;

        if (rn > 0) {
          for (let i = 1; i <= rn; i += 1) {
            height += rows.getHeight(ri + i);
          }
        }
        if (cn > 0) {
          for (let i = 1; i <= cn; i += 1) {
            width += cols.getWidth(ci + i);
          }
        }
      }
    }

    return {
      left,
      top,
      width,
      height,
      cell,
    };
  }

  getCell(ri, ci) {
    return this.rows.getCell(ri, ci);
  }

  resolveDynamicVariable(text) {
    const trigger = this.settings?.mentionProgress?.trigger;
    let resolved = false;
    let resolving = true;
    if (trigger && text?.includes?.(trigger)) {
      let regex = new RegExp(`\\${trigger}\\S*`, "g");
      const map = this?.variables?.map ?? {};
      text = String(text ?? "").replace(regex, (match) => {
        const variableName = match?.replaceAll?.(" ", "_")?.toLowerCase?.();
        if (map.hasOwnProperty(match) || map.hasOwnProperty(variableName)) {
          const { value, resolved: isResolved } =
            map[match] || map[variableName];
          resolved = isResolved;
          resolving = false;
          return value;
        } else {
          if (
            trigger === "#" &&
            EXCEL_ERRORS?.includes(
              match.replaceAll(" ", "").substring(1, match.length)
            )
          ) {
            resolved = true;
            resolving = false;
          } else {
            resolved = false;
            resolving = true;
          }
          return match;
        }
      });
    }
    return { text: text, resolved: resolved, resolving: resolving };
  }

  getCellTextOrDefault(ri, ci) {
    const cell = this.getCell(ri, ci) ?? {};
    const text = cell && cell.text ? cell.text : "";
    return this.resolveDynamicVariable.call(this, text)?.text;
  }

  getCellFormulaOrTextOrDefault(ri, ci) {
    const cell = this.getCell(ri, ci) ?? {};
    const text = cell && cell.f ? cell.f : cell.text ? cell.text : "";
    const { flipSign } = cell.cellMeta ?? {};
    const resolvedValue = this.resolveDynamicVariable.call(this, text)?.text;
    let finalValue = resolvedValue;
    const trigger = this.settings?.mentionProgress?.trigger;
    // Check if the the cell is DV cell
    if (
      flipSign &&
      trigger &&
      text?.startsWith?.(trigger) &&
      !isNaN(Number(finalValue)) &&
      resolvedValue !== ""
    ) {
      finalValue = (Number(finalValue) * -1).toString();
    }
    return finalValue;
  }

  getCellStyle(ri, ci) {
    const cell = this.getCell(ri, ci);
    if (cell && cell.style !== undefined) {
      return this.styles[cell.style];
    }
    return null;
  }

  getCellStyleOrDefault(ri, ci) {
    const { styles, rows } = this;
    const cell = rows.getCell(ri, ci);
    const cellStyle =
      cell && cell.style !== undefined ? styles[cell.style] : {};
    return helper.merge(this.defaultStyle(), cellStyle);
  }

  getCellMetaOrDefault(ri, ci) {
    const { rows } = this;
    const cell = rows.getCell(ri, ci);
    return cell?.cellMeta ?? {};
  }

  getSelectedCellStyle() {
    const { ri, ci } = this.selector;
    return this.getCellStyleOrDefault(ri, ci);
  }

  getSelectedCellMetaData() {
    const { ri, ci } = this.selector;
    return this.getCellMetaOrDefault(ri, ci);
  }

  setCellData(ri, ci, data) {
    const { rows } = this;
    // 保存原始单元格数据
    const oldCell = rows.getCell(ri, ci) || {};
    
    // 合并数据，保留特殊属性
    const cell = Object.assign({}, oldCell, data);
    
    // 确保特殊属性被保留
    if (oldCell.cellType) {
      cell.cellType = data.cellType || oldCell.cellType;
    }
    
    if (oldCell.treeData && !data.treeData) {
      cell.treeData = oldCell.treeData;
    }
    
    if (oldCell.popupData && !data.popupData) {
      cell.popupData = oldCell.popupData;
    }
    
    if (oldCell.selectedValue && !data.selectedValue) {
      cell.selectedValue = oldCell.selectedValue;
    }
    
    // 根据单元格类型设置默认值
    if (cell.cellType === 'tree' && !cell.text) {
      cell.text = '点击选择';
    } else if (cell.cellType === 'popup' && !cell.text) {
      cell.text = '点击打开';
    } else if (cell.cellType === 'date' && !cell.text) {
      cell.text = new Date().toISOString().split('T')[0];
    }
    
    // 设置单元格数据
    rows.setCell(ri, ci, cell);
    // 触发change事件
    this.change(this);
  }

  // state: input | finished
  setCellText(ri, ci, text, state = 'finished') {
    const { rows, history, validations } = this;

    if (state === "finished") {
      history.add(this.getData());
      rows.setCellText(ri, ci, text);

    } else {
      rows.setCellText(ri, ci, text);
      this.change(this.getData());

    }
    // validator
    validations.validate(ri, ci, text);
  }

  setCellProperty(ri, ci, key, value) {
    this.rows.setCellProperty(ri, ci, key, value);
  }

  freezeIsActive() {
    const [ri, ci] = this.freeze;
    return ri > 0 || ci > 0;
  }

  setFreeze(ri, ci) {
    this.changeData(() => {
      this.freeze = [ri, ci];
    });
  }

  freezeTotalWidth() {
    return this.cols.sumWidth(0, this.freeze[1]);
  }

  freezeTotalHeight() {
    return this.rows.sumHeight(0, this.freeze[0]);
  }

  setRowHeight(ri, height) {
    this.changeData(() => {
      this.rows.setHeight(ri, height);
    });
  }

  setColWidth(ci, width) {
    this.changeData(() => {
      this.cols.setWidth(ci, width);
    });
  }

  setCellMerge(ri, ci, merge) {
    if (merge?.length > 1) {
      const { rows, merges } = this;
      const cellRange = new CellRange(ri, ci, ri + merge[0], ci + merge[1]);
      this.changeData(() => {
        rows.setCellMerge(ri, ci, merge);
        merges.add(cellRange);
        const cell = rows.getCellOrNew(ri, ci);
        rows.deleteCells(cellRange);
        rows.setCell(ri, ci, cell);
      });
    }
  }

  getRowHeight(ri) {
    return this.rows.getHeight(ri);
  }

  getColWidth(ci) {
    return this.cols.getWidth(ci);
  }

  viewHeight() {
    const { view, showToolbar, showBottomBar } = this.settings;
    let h = view.height();
    if (showBottomBar) {
      h -= bottombarHeight;
    }
    if (showToolbar) {
      h -= toolbarHeight;
    }
    return h;
  }

  viewWidth() {
    return this.settings.view.width();
  }

  freezeViewRange() {
    const [ri, ci] = this.freeze;
    return new CellRange(
      0,
      0,
      ri - 1,
      ci - 1,
      this.freezeTotalWidth(),
      this.freezeTotalHeight()
    );
  }

  contentRange() {
    const { rows, cols } = this;
    const [ri, ci] = rows.maxCell();
    const h = rows.sumHeight(0, ri + 1);
    const w = cols.sumWidth(0, ci + 1);
    return new CellRange(0, 0, ri, ci, w, h);
  }

  exceptRowTotalHeight(sri, eri) {
    const { exceptRowSet, rows } = this;
    const exceptRows = Array.from(exceptRowSet);
    let exceptRowTH = 0;
    exceptRows.forEach((ri) => {
      if (ri < sri || ri > eri) {
        const height = rows.getHeight(ri);
        exceptRowTH += height;
      }
    });
    return exceptRowTH;
  }

  viewRange() {
    const {
      scroll,
      rows,
      cols,
      freeze,
      includeRowSet,
      autoFilter,
      exceptRowSet,
    } = this;
    let { ri, ci } = scroll;
    if (ri <= 0) [ri] = freeze;
    if (ci <= 0) [, ci] = freeze;

    let [x, y] = [0, 0];
    let [eri, eci] = [rows.len, cols.len];
    for (let i = ri; i < rows.len; i += 1) {
      if (autoFilter.active() && includeRowSet.size > 0) {
        if (includeRowSet.has(i)) y += rows.getHeight(i);
        else if (!exceptRowSet.has(i)) y += rows.getHeight(i);
      } else {
        y += rows.getHeight(i);
      }
      eri = i;
      if (y > this.viewHeight()) break;
    }
    for (let j = ci; j < cols.len; j += 1) {
      x += cols.getWidth(j);
      eci = j;
      if (x > this.viewWidth()) break;
    }
    return new CellRange(ri, ci, eri, eci, x, y);
  }

  eachMergesInView(viewRange, cb) {
    this.merges.filterIntersects(viewRange).forEach((it) => cb(it));
  }

  hideRowsOrCols() {
    const { rows, cols, selector } = this;
    const [rlen, clen] = selector.size();
    const { sri, sci, eri, eci } = selector.range;
    if (rlen === rows.len) {
      for (let ci = sci; ci <= eci; ci += 1) {
        cols.setHide(ci, true);
      }
    } else if (clen === cols.len) {
      for (let ri = sri; ri <= eri; ri += 1) {
        rows.setHide(ri, true);
      }
    }
  }

  // type: row | col
  // index row-index | col-index
  unhideRowsOrCols(type, index) {
    this[`${type}s`].unhide(index);
  }

  rowEach(min, max, cb) {
    let y = 0;
    const { rows } = this;
    const frset = this.exceptRowSet;
    const frary = [...frset];
    let offset = 0;
    for (let i = 0; i < frary.length; i += 1) {
      if (frary[i] < min) {
        offset += 1;
      }
    }

    for (let i = min + offset; i <= max + offset; i += 1) {
      if (frset.has(i)) {
        offset += 1;
      } else {
        const rowHeight = rows.getHeight(i);
        if (rowHeight > 0) {
          cb(i, y, rowHeight);
          y += rowHeight;
          if (y > this.viewHeight()) break;
        }
      }
    }
  }

  colEach(min, max, cb) {
    let x = 0;
    const { cols } = this;
    for (let i = min; i <= max; i += 1) {
      const colWidth = cols.getWidth(i);
      if (colWidth > 0) {
        cb(i, x, colWidth);
        x += colWidth;
        if (x > this.viewWidth()) break;
      }
    }
  }

  defaultStyle() {
    return this.settings.style;
  }

  addStyle(nstyle) {
    const { styles } = this;

    for (let i = 0; i < styles.length; i += 1) {
      const style = styles[i];
      if (helper.equals(style, nstyle)) return i;
    }
    styles.push(nstyle);
    return styles.length - 1;
  }

  changeData(cb) {
    this.history.add(this.getData());
    cb();
    this.change(this.getData());
  }

  setData(d) {
    Object.keys(d).forEach((property) => {
      if (
        property === "merges" ||
        property === "rows" ||
        property === "cols" ||
        property === "validations" ||
        property === "sheetConfig"
      ) {
        this[property].setData(d[property]);
        // 确保列数被正确设置
        if (property === "cols" && d[property] && d[property].len) {
          this.cols.len = d[property].len;
        }
      } else if (property === "freeze") {
        const [x, y] = expr2xy(d[property]);
        this.freeze = [y, x];
      } else if (property === "autofilter") {
        this.autoFilter.setData(d[property]);
      } else if (d[property] !== undefined) {
        this[property] = d[property];
      }
    });

    // 处理 sheetConfig
    if (d.sheetConfig && d.sheetConfig.settings && d.sheetConfig.settings.dataListRange) {
      // 同步到 settings 和 rows
      this.settings.dataListRange = d.sheetConfig.settings.dataListRange;
      this.rows.dataListRange = d.sheetConfig.settings.dataListRange;
    }

    return this;
  }

  getVariables() {
    const { rows, rootContext } = this;
    const trigger = rootContext?.options?.mentionProgress?.trigger;
    const rowData = rows._;
    const map = new Set();
    if (trigger) {
      let regex = new RegExp(`\\${trigger}\\S*`, "g");
      for (let rowIndex in rowData) {
        const cells = rowData[rowIndex]?.cells;
        for (let cellIndex in cells) {
          const cell = cells[cellIndex];
          const text = cell.f || cell.text;
          text?.replaceAll?.(regex, (match) => {
            if (
              trigger === "#" &&
              !EXCEL_ERRORS?.includes(match.substring(1, match.length))
            )
              map.add(match);
            else if (trigger !== "#") map.add(match);
          });
        }
      }
    }
    return map;
  }

  getData() {
    const {
      name,
      freeze,
      styles,
      merges,
      rows,
      cols,
      validations,
      autoFilter,
      sheetConfig,
      sheetId,
    } = this;
    return {
      name,
      freeze: xy2expr(freeze[1], freeze[0]),
      styles,
      merges: merges.getData(),
      rows: rows.getData(),
      cols: cols.getData(),
      validations: validations.getData(),
      autofilter: autoFilter.getData(),
      sheetConfig: sheetConfig.getData(),
      sheetId,
    };
  }

  getRootContext() {
    return this.rootContext;
  }

  updateOtherSheetsFormulaOnInsert(type, si, n = 1) {
    if (type === "row") {
      this.rootContext.datas.forEach((data) => {
        if (data.name === this.name) return;
        data.rows.each((ri, row) => {
          data.rows.eachCells(ri, (ci, cell) => {
            if (
              (cell.text && cell.text[0] === "=") ||
              (cell.text && cell.text[0] === "$")
            ) {
              cell.text = replaceCellRefWithNew(
                cell.text,
                (word) => expr2expr(word, 0, n, (x, y) => y >= si),
                {
                  isSameSheet: false,
                  sheetName: this.name,
                }
              );
            }
          });
        });
      });
    } else {
      this.rootContext.datas.forEach((data) => {
        if (data.name === this.name) return;
        data.rows.each((ri, row) => {
          data.rows.eachCells(ri, (ci, cell) => {
            if (
              (cell.text && cell.text[0] === "=") ||
              (cell.text && cell.text[0] === "$")
            ) {
              cell.text = replaceCellRefWithNew(
                cell.text,
                (word) => expr2expr(word, n, 0, (x) => x >= si),
                {
                  isSameSheet: false,
                  sheetName: this.name,
                }
              );
            }
          });
        });
      });
    }
  }

  updateOtherSheetsFormulaOnDelete(type, si, ei) {
    const n = ei - si + 1;
    if (type === "row") {
      this.rootContext.datas.forEach((data) => {
        if (data.name === this.name) return;
        data.rows.each((ri, row) => {
          data.rows.eachCells(ri, (ci, cell) => {
            if (
              (cell.text && cell.text[0] === "=") ||
              (cell.text && cell.text[0] === "$")
            ) {
              cell.text = replaceCellRefWithNew(
                cell.text,
                (word) => expr2expr(word, 0, -n, (x, y) => y > ei),
                {
                  isSameSheet: false,
                  sheetName: this.name,
                }
              );
            }
          });
        });
      });
    } else {
      this.rootContext.datas.forEach((data) => {
        if (data.name === this.name) return;
        data.rows.each((ri, row) => {
          data.rows.eachCells(ri, (ci, cell) => {
            if (
              (cell.text && cell.text[0] === "=") ||
              (cell.text && cell.text[0] === "$")
            ) {
              cell.text = replaceCellRefWithNew(
                cell.text,
                (word) => expr2expr(word, -n, 0, (x) => x > ei),
                {
                  isSameSheet: false,
                  sheetName: this.name,
                }
              );
            }
          });
        });
      });
    }
  }

  setRowId(ri, id) {
    this.changeData(() => {
      this.rows.setRowId(ri, id);
    });
  }

  getRowId(ri) {
    return this.rows.getRowId(ri);
  }

  getRowByIndex(ri) {
    return this.rows._[ri];
  }

  getRowIndexById(id) {
    return this.rows.getRowIndexById(id);
  }

  getRowDataById(id) {
    return this.rows.getRowDataById(id);
  }

  // 处理右键菜单中设置行ID的动作
  handleSetRowId() {
    const { selector } = this;
    const { sri } = selector.range;
    const currentId = this.getRowId(sri) || "";

    const modal = new RowIdModal();
    modal.show((newId) => {
      if (newId !== null && newId !== "") {
        this.setRowId(sri, newId);
      }
    }, currentId); // 传入当前ID作为默认值
  }

  // 检查单元格是否可编辑
  canEditCell(ri, ci) {
    const cell = this.getCell(ri, ci);
    const mode = this.settings?.mode;
    
    // 在设计模式下，所有单元格都可编辑
    if (mode === 'design') {
      return true;
    }
    
    // 在其他模式下，遵循单元格的nonEditable属性
    return !cell || !cell.nonEditable;
  }

  getValidator(ri, ci) {
    return this.validations.get(ri, ci); // 或类似逻辑
  }
}
