import helper from "./helper";
import { expr2expr } from "./alphabet";
import { replaceCellRefWithNew } from "../utils";

class Rows {
  constructor(dataProxyContext, { len, height, minHeight }, options = {}) {
    this.options = options;
    this._ = {};
    this.len = len;
    // default row height
    this.height = height;
    this.data = dataProxyContext;
    this.isHeightChanged = false;
    this.minHeight = minHeight ?? height;
  }

  getHeight(ri) {
    if (this.isHide(ri)) return 0;
    const row = this.get(ri);
    if (row && row.height) {
      return row.height;
    }
    return this.height;
  }

  setHeight(ri, v, isTextWrap = false) {
    const row = this.getOrNew(ri);
    row.height = v;
    if (!isTextWrap) this.isHeightChanged = true;
  }

  unhide(idx) {
    let index = idx;
    while (index > 0) {
      index -= 1;
      if (this.isHide(index)) {
        this.setHide(index, false);
      } else break;
    }
  }

  isHide(ri) {
    const row = this.get(ri);
    return row && row.hide;
  }

  setHide(ri, v) {
    const row = this.getOrNew(ri);
    if (v === true) row.hide = true;
    else delete row.hide;
  }

  setStyle(ri, style) {
    const row = this.getOrNew(ri);
    row.style = style;
  }

  sumHeight(min, max, exceptSet) {
    return helper.rangeSum(min, max, (i) => {
      if (exceptSet && exceptSet.has(i)) return 0;
      return this.getHeight(i);
    });
  }

  totalHeight() {
    return this.sumHeight(0, this.len);
  }

  get(ri) {
    return this._[ri];
  }

  getOrNew(ri) {
    this._[ri] = this._[ri] || { cells: {} };
    if (ri >= this.len) this.len += 1;
    return this._[ri];
  }

  getCell(ri, ci) {
    const row = this.get(ri);
    if (
      row !== undefined &&
      row.cells !== undefined &&
      row.cells[ci] !== undefined
    ) {
      return row.cells[ci];
    }
    return null;
  }

  getCellMerge(ri, ci) {
    const cell = this.getCell(ri, ci);
    if (cell && cell.merge) return cell.merge;
    return [0, 0];
  }

  getCellOrNew(ri, ci) {
    const row = this.getOrNew(ri);
    row.cells[ci] = row.cells[ci] || {};
    if (ci >= this?.data?.cols?.len) this.data.cols.len += 1;
    return row.cells[ci];
  }

  // what: all | text | format
  setCell(ri, ci, cell, what = "all") {
    const row = this.getOrNew(ri);
    if (what === "all") {
      row.cells[ci] = cell;
    } else if (what === "text") {
      row.cells[ci] = row.cells[ci] || {};
      row.cells[ci].text = cell.text;
    } else if (what === "format") {
      row.cells[ci] = row.cells[ci] || {};
      row.cells[ci].style = cell.style;
      if (cell.merge) row.cells[ci].merge = cell.merge;
    }
  }

  setCellText(ri, ci, text) {
    // console.log("setCellText:", ri, ci, text);
    const cell = this.getCellOrNew(ri, ci);
    if (cell.editable !== false) {
      // console.log("setCellText-editable:", ri, ci, text);
      const valueSetter = this.options.valueSetter;
      if (valueSetter) {
        // console.log("setCellText-valueSetter:", ri, ci, text);
        const result = valueSetter({ ...this, text, cell });
        let retrievedText, formattedText, html;
        if (result && typeof result === "object" && !Array.isArray(result)) {
          // console.log("setCellText-valueSetter-result:", ri, ci, text);
          ({ retrievedText, formattedText, html } = result);
        } else if (Array.isArray(result)) {
          // console.log("setCellText-valueSetter-result-array:", ri, ci, text);
          [retrievedText, formattedText, html] = result;
        } else {
          // console.log("setCellText-valueSetter-result-else:", ri, ci, text);
          retrievedText = result;
          formattedText = "";
          html = "";
        }
        cell.text = retrievedText ?? "";
        cell.w = formattedText ?? "";
        cell.f = "";
        cell.h = html ?? "";
      } else {
        // console.log("setCellText-else:", ri, ci, text);
        cell.text = text;
        cell.f = "";
        cell.h = "";
      }
    }
  }

  setCellProperty(ri, ci, key, value) {
    const cell = this.getCellOrNew(ri, ci);
    if (key) {
      cell[key] = value;
    }
  }

  setCellMerge(ri, ci, merge) {
    const cell = this.getCellOrNew(ri, ci);
    cell.merge = [merge[0], merge[1]];
  }

  // what: all | format | text
  copyPaste(srcCellRange, dstCellRange, what, autofill = false, cb = () => {}) {
    const { sri, sci, eri, eci } = srcCellRange;
    const dsri = dstCellRange.sri;
    const dsci = dstCellRange.sci;
    const deri = dstCellRange.eri;
    const deci = dstCellRange.eci;
    const [rn, cn] = srcCellRange.size();
    const [drn, dcn] = dstCellRange.size();
    // console.log(srcIndexes, dstIndexes);
    let isAdd = true;
    let dn = 0;
    if (deri < sri || deci < sci) {
      isAdd = false;
      if (deri < sri) dn = drn;
      else dn = dcn;
    }
    for (let i = sri; i <= eri; i += 1) {
      if (this._[i]) {
        for (let j = sci; j <= eci; j += 1) {
          if (this._[i].cells && this._[i].cells[j]) {
            for (let ii = dsri; ii <= deri; ii += rn) {
              for (let jj = dsci; jj <= deci; jj += cn) {
                const nri = ii + (i - sri);
                const nci = jj + (j - sci);
                const ncell = helper.cloneDeep(this._[i].cells[j]);
                // ncell.text
                if (autofill && ncell && ncell.text && ncell.text.length > 0) {
                  const { text } = ncell;
                  let n = jj - dsci + (ii - dsri) + 2;
                  if (!isAdd) {
                    n -= dn + 1;
                  }
                  if (text[0] === "=") {
                    ncell.text = text.replace(/[a-zA-Z]{1,3}\d+/g, (word) => {
                      let [xn, yn] = [0, 0];
                      if (sri === dsri) {
                        xn = n - 1;
                        // if (isAdd) xn -= 1;
                      } else {
                        yn = n - 1;
                      }
                      if (/^\d+$/.test(word)) return word;
                      return expr2expr(word, xn, yn);
                    });
                  } else if (
                    (rn <= 1 && cn > 1 && (dsri > eri || deri < sri)) ||
                    (cn <= 1 && rn > 1 && (dsci > eci || deci < sci)) ||
                    (rn <= 1 && cn <= 1)
                  ) {
                    const result = /[\\.\d]+$/.exec(text);
                    // console.log('result:', result);
                    if (result !== null) {
                      const index = Number(result[0]) + n - 1;
                      ncell.text = text.substring(0, result.index) + index;
                    }
                  }
                }
                this.setCell(nri, nci, ncell, what);
                cb(nri, nci, ncell);
              }
            }
          }
        }
      }
    }
  }

  cutPaste(srcCellRange, dstCellRange) {
    const ncellmm = {};
    this.each((ri) => {
      this.eachCells(ri, (ci) => {
        let nri = parseInt(ri, 10);
        let nci = parseInt(ci, 10);
        if (srcCellRange.includes(ri, ci)) {
          nri = dstCellRange.sri + (nri - srcCellRange.sri);
          nci = dstCellRange.sci + (nci - srcCellRange.sci);
        }
        ncellmm[nri] = ncellmm[nri] || { cells: {} };
        ncellmm[nri].cells[nci] = this._[ri].cells[ci];
      });
    });
    this._ = ncellmm;
  }

  // src: Array<Array<String>>
  paste(src, dstCellRange) {
    if (src.length <= 0) return;
    const { sri, sci } = dstCellRange;
    src.forEach((row, i) => {
      const ri = sri + i;
      row.forEach((cell, j) => {
        const ci = sci + j;
        this.setCellText(ri, ci, cell);
      });
    });
  }

  insert(sri, n = 1) {
    const ndata = {};
    this.each((ri, row) => {
      let nri = parseInt(ri, 10);
      if (nri >= sri) {
        nri += n;
        this.eachCells(ri, (ci, cell) => {
          if (cell.text && cell.text[0] === "=" || cell.text && cell.text[0] === "$") {
            cell.text = replaceCellRefWithNew(
              cell.text,
              (word) => expr2expr(word, 0, n, (x, y) => y >= sri),
              {
                sheetName: this.data.name,
                isSameSheet: true,
              }
            );
          }
        });
      }
      ndata[nri] = row;
    });
    this._ = ndata;
    this.len += n;
  }

  delete(sri, eri) {
    const n = eri - sri + 1;
    const ndata = {};
    this.each((ri, row) => {
      const nri = parseInt(ri, 10);
      if (nri < sri) {
        ndata[nri] = row;
      } else if (ri > eri) {
        ndata[nri - n] = row;
        this.eachCells(ri, (ci, cell) => {
          if (cell.text && cell.text[0] === "=" || cell.text && cell.text[0] === "$") {
            cell.text = replaceCellRefWithNew(
              cell.text,
              (word) => expr2expr(word, 0, -n, (x, y) => y > eri),
              {
                sheetName: this.data.name,
                isSameSheet: true,
              }
            );
          }
        });
      }
    });
    this._ = ndata;
    this.len -= n;
  }

  insertColumn(sci, n = 1) {
    this.each((ri, row) => {
      const rndata = {};
      this.eachCells(ri, (ci, cell) => {
        let nci = parseInt(ci, 10);
        if (nci >= sci) {
          nci += n;
          if (cell.text && cell.text[0] === "=" || cell.text && cell.text[0] === "$") {
            cell.text = replaceCellRefWithNew(
              cell.text,
              (word) => expr2expr(word, n, 0, (x) => x >= sci),
              {
                sheetName: this.data.name,
                isSameSheet: true,
              }
            );
          }
        }
        rndata[nci] = cell;
      });
      row.cells = rndata;
    });
  }

  deleteColumn(sci, eci) {
    const n = eci - sci + 1;
    this.each((ri, row) => {
      const rndata = {};
      this.eachCells(ri, (ci, cell) => {
        const nci = parseInt(ci, 10);
        if (nci < sci) {
          rndata[nci] = cell;
        } else if (nci > eci) {
          rndata[nci - n] = cell;
          if (cell.text && cell.text[0] === "=" || cell.text && cell.text[0] === "$") {
            cell.text = replaceCellRefWithNew(
              cell.text,
              (word) => expr2expr(word, -n, 0, (x) => x > eci),
              {
                sheetName: this.data.name,
                isSameSheet: true,
              }
            );
          }
        }
      });
      row.cells = rndata;
    });
  }

  // what: all | text | format | merge
  deleteCells(cellRange, what = "all") {
    cellRange.each((i, j) => {
      this.deleteCell(i, j, what);
    });
  }

  // what: all | text | format | merge
  deleteCell(ri, ci, what = "all") {
    const row = this.get(ri);
    if (row !== null) {
      const cell = this.getCell(ri, ci);
      if (cell !== null && cell.editable !== false) {
        if (what === "all") {
          delete row.cells[ci];
        } else if (what === "text") {
          if (cell.text) delete cell.text;
          if (cell.value) delete cell.value;
          if (cell.h) cell.h = "";
          if (cell.w) cell.w = "";
          if (cell.f) cell.f = "";
        } else if (what === "format") {
          if (cell.style !== undefined) delete cell.style;
          if (cell.merge) delete cell.merge;
        } else if (what === "merge") {
          if (cell.merge) delete cell.merge;
        }
      }
    }
  }

  maxCell() {
    const keys = Object.keys(this._);
    const ri = keys[keys.length - 1];
    const col = this._[ri];
    if (col) {
      const { cells } = col;
      const ks = Object.keys(cells);
      const ci = ks[ks.length - 1];
      return [parseInt(ri, 10), parseInt(ci, 10)];
    }
    return [0, 0];
  }

  each(cb) {
    Object.entries(this._).forEach(([ri, row]) => {
      cb(ri, row);
    });
  }

  eachCells(ri, cb) {
    if (this._[ri] && this._[ri].cells) {
      Object.entries(this._[ri].cells).forEach(([ci, cell]) => {
        cb(ci, cell);
      });
    }
  }

  setData(d) {
    if (d.len) {
      this.len = d.len;
      delete d.len;
    }
    this._ = d;
  }

  // 添加通过id获取行号的方法
  getRowIndexById(id) {
    for (let ri in this._) {
      if (this._[ri].id === id) {
        return parseInt(ri);
      }
    }
    return -1;
  }

  // 添加设置行id的方法
  setRowId(ri, id) {
    const row = this.getOrNew(ri);
    if (!row.id) {
      row.id = id;
    } else {
      row.id = id;
    }
  }

  // 添加获取行id的方法
  getRowId(ri) {
    const row = this._[ri];
    return row ? row.id : null;
  }

  getData() {
    const data = {};
    Object.entries(this._).forEach(([ri, row]) => {
      if (row.id || Object.keys(row.cells).length > 0 || row.height !== this.height) {
        data[ri] = {
          cells: row.cells,
          height: row.height,
          id: row.id
        };
      }
    });
    data.len = Object.keys(data).length;
    return data;
  }

  setData(data) {
    if (data.len) {
      this.len = data.len;
    }
    Object.entries(data).forEach(([ri, row]) => {
      if (ri !== 'len') {
        const r = this.getOrNew(parseInt(ri, 10));
        r.cells = row.cells || {};
        if (row.height) r.height = row.height;
        if (row.id) r.id = row.id; // 从导入数据中恢复id
      }
    });
  }
}

export default {};
export { Rows };
