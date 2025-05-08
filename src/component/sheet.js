/* global window */
import { h } from "./element";
import { bind, mouseMoveUp, bindTouch, createEventEmitter } from "./event";
import Resizer from "./resizer";
import Scrollbar from "./scrollbar";
import Selector from "./selector";
import Editor from "./editor";
import Print from "./print";
import ContextMenu from "./contextmenu";
import Table from "./table";
import Toolbar from "./toolbar/index";
import ModalValidation from "./modal_validation";
import SortFilter from "./sort_filter";
import { xtoast } from "./message";
import { cssPrefix } from "../config";
import { formulas } from "../core/formula";
import { constructFormula, getCellName } from "../algorithm/cellInjection";
import Comment from "./comment";
import { CELL_REF_REGEX, SHEET_TO_CELL_REF_REGEX } from "../constants";
import { expr2xy } from "../core/alphabet";
import Datepicker from './datepicker';

/**
 * @desc throttle fn
 * @param func function
 * @param wait Delay in milliseconds
 */
function throttle(func, wait) {
  let timeout;
  return (...arg) => {
    const that = this;
    const args = arg;
    if (!timeout) {
      timeout = setTimeout(() => {
        timeout = null;
        func.apply(that, args);
      }, wait);
    }
  };
}

function debounce(func, wait, immediate = false) {
  let timeout;
  return function executedFunction(...args) {
    const context = this;
    const later = function() {
      timeout = null;
      if (!immediate) func.apply(context, args);
    };
    const callNow = immediate && !timeout;
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
    if (callNow) func.apply(context, args);
  };
}

function scrollbarMove() {
  const { data, verticalScrollbar, horizontalScrollbar } = this;
  const { l, t, left, top, width, height } = data.getSelectedRect();
  const tableOffset = this.getTableOffset();
  if (Math.abs(left) + width > tableOffset.width) {
    horizontalScrollbar.move({ left: l + width - tableOffset.width });
  } else {
    const fsw = data.freezeTotalWidth();
    if (left < fsw) {
      horizontalScrollbar.move({ left: l - 1 - fsw });
    }
  }
  if (Math.abs(top) + height > tableOffset.height) {
    verticalScrollbar.move({ top: t + height - tableOffset.height - 1 });
  } else {
    const fsh = data.freezeTotalHeight();
    if (top < fsh) {
      verticalScrollbar.move({ top: t - 1 - fsh });
    }
  }
}

function formulaProgress() {
  if (!this.editor.formulaCell) return;
  const single = this.data.isSingleSelected();
  const text = this.editor.inputText;
  if (single) {
    const { ri, ci } = this.data.selector;
    const cellName = getCellName(ri, ci);
    const updatedText = constructFormula(text, cellName, false);
    this.editor.setText(updatedText);
  } else {
    const { sri, eri, sci, eci } = this.selector.range;
    const startCell = getCellName(sri, sci);
    const endCell = getCellName(eri, eci);
    const range = `${startCell}:${endCell}`;
    const updatedText = constructFormula(text, range, true);
    this.editor.setText(updatedText);
  }
}

function selectorSet(multiple, ri, ci, indexesUpdated = true, moving = false) {
  const { table, selector, toolbar, data, contextMenu, editor } = this;
  const cell = data.getCell(ri, ci);
  if (multiple) {
    selector.setEnd(ri, ci, moving);
    if (!moving) {
      if (!editor.formulaCell) {
        this.trigger("cells-selected", cell, selector.range);
      }
      formulaProgress.call(this);
    }
  } else {
    selector.set(ri, ci, indexesUpdated);
    if (!moving && ri !== -1 && ci !== -1) {
      if (!editor.formulaCell) {
        this.trigger("cell-selected", cell, ri, ci);
      }
      formulaProgress.call(this);
    }
  }
  contextMenu.setMode(ri === -1 || ci === -1 ? "row-col" : "range");
  toolbar.reset();
  table.render();
}

// multiple: boolean
// direction: left | right | up | down | row-first | row-last | col-first | col-last
function selectorMove(multiple, direction, isMoving = true) {
  const { selector, data } = this;
  const { rows, cols } = data;
  let [ri, ci] = selector.indexes;
  const { eri, eci } = selector.range;
  let isLastOrFirst = false;
  if (multiple) {
    [ri, ci] = selector.moveIndexes;
  }
  if (direction === "left") {
    if (ci > 0) ci -= 1;
  } else if (direction === "right") {
    if (eci !== ci) ci = eci;
    if (ci < cols.len - 1) ci += 1;
  } else if (direction === "up") {
    if (ri > 0) ri -= 1;
  } else if (direction === "down") {
    if (eri !== ri) ri = eri;
    if (ri < rows.len - 1) ri += 1;
  } else if (direction === "row-first") {
    ci = 0;
    isLastOrFirst = true;
  } else if (direction === "row-last") {
    ci = cols.len - 1;
    isLastOrFirst = true;
  } else if (direction === "col-first") {
    ri = 0;
    isLastOrFirst = true;
  } else if (direction === "col-last") {
    ri = rows.len - 1;
    isLastOrFirst = true;
  }
  if (multiple) {
    selector.moveIndexes = [ri, ci];
  }
  selectorSet.call(this, multiple, ri, ci, true, !isLastOrFirst && isMoving);
  scrollbarMove.call(this);
}

// private methods
function overlayerMousemove(evt) {
  if (evt.buttons !== 0) return;
  if (evt.target.className === `${cssPrefix}-resizer-hover`) return;
  const { offsetX, offsetY } = evt;
  const { rowResizer, colResizer, tableEl, data } = this;
  const { rows, cols } = data;
  if (offsetX > cols.indexWidth && offsetY > rows.height) {
    rowResizer.hide();
    colResizer.hide();
    return;
  }
  const tRect = tableEl.box();
  const cRect = data.getCellRectByXY(evt.offsetX, evt.offsetY);
  if (cRect.ri >= 0 && cRect.ci === -1) {
    cRect.width = cols.indexWidth;
    rowResizer.show(cRect, {
      width: tRect.width,
    });
    if (rows.isHide(cRect.ri - 1)) {
      rowResizer.showUnhide(cRect.ri);
    } else {
      rowResizer.hideUnhide();
    }
  } else {
    rowResizer.hide();
  }
  if (cRect.ri === -1 && cRect.ci >= 0) {
    cRect.height = rows.height;
    colResizer.show(cRect, {
      height: tRect.height,
    });
    if (cols.isHide(cRect.ci - 1)) {
      colResizer.showUnhide(cRect.ci);
    } else {
      colResizer.hideUnhide();
    }
  } else {
    colResizer.hide();
  }
}

// let scrollThreshold = 15;
function overlayerMousescroll(evt) {
  // scrollThreshold -= 1;
  // if (scrollThreshold > 0) return;
  // scrollThreshold = 15;

  const { verticalScrollbar, horizontalScrollbar, data } = this;
  const { top } = verticalScrollbar.scroll();
  const { left } = horizontalScrollbar.scroll();

  const { rows, cols } = data;

  // deltaY for vertical delta
  const { deltaY, deltaX } = evt;
  const loopValue = (ii, vFunc) => {
    let i = ii;
    let v = 0;
    do {
      v = vFunc(i);
      i += 1;
    } while (v <= 0);
    return v;
  };
  // if (evt.detail) deltaY = evt.detail * 40;
  const moveY = (vertical) => {
    if (vertical > 0) {
      // up
      const ri = data.scroll.ri + 1;
      if (ri < rows.len) {
        const rh = loopValue(ri, (i) => rows.getHeight(i));
        verticalScrollbar.move({ top: top + rh - 1 });
      }
    } else {
      // down
      const ri = data.scroll.ri - 1;
      if (ri >= 0) {
        const rh = loopValue(ri, (i) => rows.getHeight(i));
        verticalScrollbar.move({ top: ri === 0 ? 0 : top - rh });
      }
    }
  };

  // deltaX for Mac horizontal scroll
  const moveX = (horizontal) => {
    if (horizontal > 0) {
      // left
      const ci = data.scroll.ci + 1;
      if (ci < cols.len) {
        const cw = loopValue(ci, (i) => cols.getWidth(i));
        horizontalScrollbar.move({ left: left + cw - 1 });
      }
    } else {
      // right
      const ci = data.scroll.ci - 1;
      if (ci >= 0) {
        const cw = loopValue(ci, (i) => cols.getWidth(i));
        horizontalScrollbar.move({ left: ci === 0 ? 0 : left - cw });
      }
    }
  };
  const absDeltaY = Math.abs(deltaY);
  const absDeltaX = Math.abs(deltaX);

  // detail for windows/mac firefox vertical scroll
  if (/Firefox/i.test(window.navigator.userAgent))
    throttle(moveY(evt.detail), 50);
  if (absDeltaY > absDeltaX && absDeltaX == 0) throttle(moveY(deltaY), 50);
  if (absDeltaX > absDeltaY && absDeltaY == 0) throttle(moveX(deltaX), 50);
}

function overlayerTouch(direction, distance) {
  const { verticalScrollbar, horizontalScrollbar } = this;
  const { top } = verticalScrollbar.scroll();
  const { left } = horizontalScrollbar.scroll();

  if (direction === "left" || direction === "right") {
    horizontalScrollbar.move({ left: left - distance });
  } else if (direction === "up" || direction === "down") {
    verticalScrollbar.move({ top: top - distance });
  }
}

function verticalScrollbarSet() {
  const { data, verticalScrollbar } = this;
  const { height } = this.getTableOffset();
  const erth = data.exceptRowTotalHeight(0, -1);
  verticalScrollbar.set(height, data.rows.totalHeight() - erth);
}

function horizontalScrollbarSet() {
  const { data, horizontalScrollbar } = this;
  const { width } = this.getTableOffset();
  if (data) {
    horizontalScrollbar.set(width, data.cols.totalWidth());
  }
}

function sheetFreeze() {
  const { selector, data, editor } = this;
  const [ri, ci] = data.freeze;
  if (ri > 0 || ci > 0) {
    const fwidth = data.freezeTotalWidth();
    const fheight = data.freezeTotalHeight();
    editor.setFreezeLengths(fwidth, fheight);
  }
  selector.resetAreaOffset();
}

function sheetReset() {
  const { tableEl, overlayerEl, overlayerCEl, table, toolbar, selector, el } =
    this;
  const tOffset = this.getTableOffset();
  const vRect = this.getRect();
  tableEl.attr(vRect);
  overlayerEl.offset(vRect);
  overlayerCEl.offset(tOffset);
  el.css("width", `${vRect.width}px`);
  verticalScrollbarSet.call(this);
  horizontalScrollbarSet.call(this);
  sheetFreeze.call(this);
  table.render();
  toolbar.reset();
  selector.reset();
}

function clearClipboard() {
  const { data, selector } = this;
  data.clearClipboard();
  selector.hideClipboard();
}

function copy(evt) {
  const { data, selector } = this;
  if (data.settings.mode === "read" || data.settings.mode === "preview" || data.settings.mode === "enabled") return;
  data.copy();
  data.copyToSystemClipboard(evt);
  const copiedCellRange = selector.range;
  this.trigger("copied-clipboard", copiedCellRange);
  selector.showClipboard();
}

function cut() {
  const { data, selector } = this;
  if (data.settings.mode === "read" || data.settings.mode === "preview" || data.settings.mode === "enabled") return;
  data.cut();
  const cutCellRange = selector.range;
  this.trigger("cut-clipboard", cutCellRange);
  selector.showClipboard();
}

function paste(what, evt) {
  const { data } = this;
  const eventTrigger = (type = "system") => {
    this.trigger("pasted-clipboard", { type, data });
  };
  if (data.settings.mode === "read") return;
  if (data.clipboard.isClear()) {
    const resetSheet = () => sheetReset.call(this);
    // pastFromSystemClipboard is async operation, need to tell it how to reset sheet and trigger event after it finishes
    // pasting content from system clipboard
    data.pasteFromSystemClipboard(evt, resetSheet);
    eventTrigger("system");
    resetSheet();
  } else if (data.paste(what, (msg) => xtoast("Tip", msg))) {
    eventTrigger("internal");
    sheetReset.call(this);
  } else if (evt) {
    const cdata = evt.clipboardData.getData("text/plain");
    this.data.pasteFromText(cdata);
    eventTrigger("system");
    sheetReset.call(this);
  }
}

function hideRowsOrCols() {
  this.data.hideRowsOrCols();
  sheetReset.call(this);
}

function unhideRowsOrCols(type, index) {
  this.data.unhideRowsOrCols(type, index);
  sheetReset.call(this);
}

function autofilter() {
  const { data } = this;
  data.autofilter();
  sheetReset.call(this);
}

function toolbarChangePaintformatPaste() {
  const { toolbar } = this;
  if (toolbar.paintformatActive()) {
    paste.call(this, "format");
    clearClipboard.call(this);
    toolbar.paintformatToggle();
  }
}

function overlayerMousedown(evt) {
  const { selector, data, table, sortFilter } = this;
  const { offsetX, offsetY } = evt;
  const isAutofillEl = evt.target.className === `${cssPrefix}-selector-corner`;
  const cellRect = data.getCellRectByXY(offsetX, offsetY);
  const { left, top, width, height } = cellRect;
  let { ri, ci } = cellRect;
  
  // 行标题区域点击 (ri >= 0 && ci === -1)
  if (ri >= 0 && ci === -1) {
    // 直接设置选择范围为整行，跳过标准选择流程，避免合并单元格影响
    const { cols } = data;
    const lastColIndex = cols.len - 1;
    
    // 初始化选择范围
    selector.range.sri = ri;
    selector.range.eri = ri;
    selector.range.sci = 0;
    selector.range.eci = lastColIndex;
    
    // 更新选择器索引
    selector.indexes = [ri, 0];
    selector.moveIndexes = [ri, lastColIndex];
    
    // 强制渲染选择器
    selector.resetAreaOffset();
    table.render();
    
    this.trigger("cells-selected", null, selector.range);
    
    // 标记当前正在进行行选择模式
    const initialRowIndex = ri;
    
    // 缓存上一次的选择范围，用于判断是否需要重新渲染
    let lastSri = ri;
    let lastEri = ri;
    
    // 使用节流函数限制更新频率
    const updateSelection = throttle(function(targetRowIndex) {
      // 只有当选择范围发生变化时才更新
      let newSri, newEri;
      
      if (targetRowIndex >= initialRowIndex) {
        // 向下拖动
        newSri = initialRowIndex;
        newEri = targetRowIndex;
      } else {
        // 向上拖动
        newSri = targetRowIndex;
        newEri = initialRowIndex;
      }
      
      // 如果选择范围没有变化，则不更新
      if (newSri === lastSri && newEri === lastEri) {
        return;
      }
      
      // 更新缓存的范围
      lastSri = newSri;
      lastEri = newEri;
      
      // 更新选择器范围
      selector.range.sri = newSri;
      selector.range.eri = newEri;
      selector.range.sci = 0;
      selector.range.eci = lastColIndex;
      
      // 更新选择器索引
      selector.moveIndexes = [targetRowIndex, lastColIndex];
      
      // 强制渲染选择器
      selector.resetAreaOffset();
      table.render();
    }, 30); // 30ms的节流间隔
    
    // 鼠标移动处理，支持多行选择
    mouseMoveUp(
      window,
      (e) => {
        // 获取当前鼠标位置下的单元格
        const rect = data.getCellRectByXY(e.offsetX, e.offsetY);
        let targetRowIndex = rect.ri;
        
        // 如果鼠标移出表格区域或在无效位置，尝试估算最近的行
        if (targetRowIndex < 0) {
          // 根据鼠标Y坐标估算行索引
          if (e.offsetY <= data.rows.height) {
            // 如果鼠标在表头上方，选择第一行
            targetRowIndex = 0;
          } else {
            // 如果鼠标低于表头，选择最后一行
            const lastVisibleRow = data.scroll.ri + data.scroll.rn;
            targetRowIndex = Math.min(lastVisibleRow, data.rows.len - 1);
          }
        }
        
        // 使用节流函数更新选择范围
        if (targetRowIndex >= 0) {
          updateSelection(targetRowIndex);
        }
      },
      () => {
        // 选择完成后触发事件
        this.trigger("cells-selected", null, selector.range);
      }
    );
    return;
  }
  
  // 列标题区域点击 (ri === -1 && ci >= 0)
  if (ri === -1 && ci >= 0) {
    // 直接设置选择范围为整列，跳过标准选择流程，避免合并单元格影响
    const { rows } = data;
    const lastRowIndex = rows.len - 1;
    
    // 初始化选择范围
    selector.range.sri = 0;
    selector.range.eri = lastRowIndex;
    selector.range.sci = ci;
    selector.range.eci = ci;
    
    // 更新选择器索引
    selector.indexes = [0, ci];
    selector.moveIndexes = [lastRowIndex, ci];
    
    // 强制渲染选择器
    selector.resetAreaOffset();
    table.render();
    
    this.trigger("cells-selected", null, selector.range);
    
    // 标记当前正在进行列选择模式
    const initialColIndex = ci;
    
    // 缓存上一次的选择范围，用于判断是否需要重新渲染
    let lastSci = ci;
    let lastEci = ci;
    
    // 使用节流函数限制更新频率
    const updateSelection = throttle(function(targetColIndex) {
      // 只有当选择范围发生变化时才更新
      let newSci, newEci;
      
      if (targetColIndex >= initialColIndex) {
        // 向右拖动
        newSci = initialColIndex;
        newEci = targetColIndex;
      } else {
        // 向左拖动
        newSci = targetColIndex;
        newEci = initialColIndex;
      }
      
      // 如果选择范围没有变化，则不更新
      if (newSci === lastSci && newEci === lastEci) {
        return;
      }
      
      // 更新缓存的范围
      lastSci = newSci;
      lastEci = newEci;
      
      // 更新选择器范围
      selector.range.sci = newSci;
      selector.range.eci = newEci;
      selector.range.sri = 0;
      selector.range.eri = lastRowIndex;
      
      // 更新选择器索引
      selector.moveIndexes = [lastRowIndex, targetColIndex];
      
      // 强制渲染选择器
      selector.resetAreaOffset();
      table.render();
    }, 30); // 30ms的节流间隔
    
    // 鼠标移动处理，支持多列选择
    mouseMoveUp(
      window,
      (e) => {
        // 获取当前鼠标位置下的单元格
        const rect = data.getCellRectByXY(e.offsetX, e.offsetY);
        let targetColIndex = rect.ci;
        
        // 如果鼠标移出表格区域或在无效位置，尝试估算最近的列
        if (targetColIndex < 0) {
          // 根据鼠标X坐标估算列索引
          if (e.offsetX <= data.cols.indexWidth) {
            // 如果鼠标在行号左侧，选择第一列
            targetColIndex = 0;
          } else {
            // 如果鼠标在行号右侧，选择最后一列
            const lastVisibleCol = data.scroll.ci + data.scroll.cn;
            targetColIndex = Math.min(lastVisibleCol, data.cols.len - 1);
          }
        }
        
        // 使用节流函数更新选择范围
        if (targetColIndex >= 0) {
          updateSelection(targetColIndex);
        }
      },
      () => {
        // 选择完成后触发事件
        this.trigger("cells-selected", null, selector.range);
      }
    );
    return;
  }
  
  // 左上角点击 (ri === -1 && ci === -1)
  if (ri === -1 && ci === -1) {
    // 直接设置选择范围为整个表格，跳过标准选择流程，避免合并单元格影响
    const { rows, cols } = data;
    const lastRowIndex = rows.len - 1;
    const lastColIndex = cols.len - 1;
    
    // 设置选择范围
    selector.range.sri = 0;
    selector.range.eri = lastRowIndex;
    selector.range.sci = 0;
    selector.range.eci = lastColIndex;
    
    // 更新选择器索引
    selector.indexes = [0, 0];
    selector.moveIndexes = [lastRowIndex, lastColIndex];
    
    // 强制渲染选择器
    selector.resetAreaOffset();
    table.render();
    
    this.trigger("cells-selected", null, selector.range);
    return;
  }
  
  // sort or filter
  const { autoFilter } = data;
  if (autoFilter.includes(ri, ci)) {
    if (left + width - 20 < offsetX && top + height - 20 < offsetY) {
      const items = autoFilter.items(ci, (r, c) => data.rows.getCell(r, c));
      sortFilter.hide();
      sortFilter.set(
        ci,
        items,
        autoFilter.getFilter(ci),
        autoFilter.getSort(ci)
      );
      sortFilter.setOffset({ left, top: top + height + 2 });
      return;
    }
  }

  if (!evt.shiftKey) {
    if (isAutofillEl) {
      selector.showAutofill(ri, ci);
    } else {
      selectorSet.call(this, false, ri, ci, true, true);
    }

    // mouse move up
    mouseMoveUp(
      window,
      (e) => {
        ({ ri, ci } = data.getCellRectByXY(e.offsetX, e.offsetY));
        if (isAutofillEl) {
          selector.showAutofill(ri, ci);
        } else if (e.buttons === 1 && !e.shiftKey) {
          selectorSet.call(this, true, ri, ci, true, true);
        }
      },
      (e) => {
        if (isAutofillEl && selector.arange && data.settings.mode !== "read") {
          if (
            data.autofill(selector.arange, "all", (msg) => xtoast("Tip", msg))
          ) {
            table.render();
          }
        } else if (!e.shiftKey) {
          const range = this.selector.range;
          const isSingleSelect =
            range.sri === range.eri && range.sci === range.eci;
          selectorSet.call(this, !isSingleSelect, ri, ci, true, false);
        }
        selector.hideAutofill();
        toolbarChangePaintformatPaste.call(this);
      }
    );
  }

  if (!isAutofillEl && evt.buttons === 1) {
    if (evt.shiftKey) {
      selectorSet.call(this, true, ri, ci);
    }
  }
}

function editorSetOffset() {
  const { editor, data } = this;
  const sOffset = data.getSelectedRect();
  const tOffset = this.getTableOffset();
  let sPosition = "top";
  if (sOffset.top > tOffset.height / 2) {
    sPosition = "bottom";
  }
  editor.setOffset(sOffset, sPosition);
}

function editorSet() {
  const { editor, data } = this;
  const selectedCell = data.getSelectedCell(); // 获取一次，避免重复调用
  
  // Existing editorSet logic
  if (editor.formulaCell) {
    return;
  }

  if (data.settings.mode === "preview" || data.settings.mode === "normal" || data.settings.mode === "enabled") {
    return;
  }
  
  // 在预览模式下检查单元格是否可编辑
  if (selectedCell && selectedCell.editable === false) {
    return;
  }
  
  editorSetOffset.call(this);
  
  // 如果cell为null，我们需要创建一个新的单元格
  if (!selectedCell) {
    const { ri, ci } = data.selector;
    data.rows.getCellOrNew(ri, ci);
  }
  
  editor.setCell(data.getSelectedCell() || {}, data.getSelectedValidator());
  clearClipboard.call(this);
}

function verticalScrollbarMove(distance) {
  const { data, table, selector } = this;
  data.scrolly(distance, () => {
    selector.resetBRLAreaOffset();
    editorSetOffset.call(this);
    table.render();
  });
}

function horizontalScrollbarMove(distance) {
  const { data, table, selector } = this;
  data.scrollx(distance, () => {
    selector.resetBRTAreaOffset();
    editorSetOffset.call(this);
    table.render();
  });
}

function rowResizerFinished(cRect, distance) {
  const { ri } = cRect;
  const { table, selector, data } = this;
  const { sri, eri } = selector.range;
  if (ri >= sri && ri <= eri) {
    for (let row = sri; row <= eri; row += 1) {
      data.rows.setHeight(row, distance);
    }
  } else {
    data.rows.setHeight(ri, distance);
  }

  table.render();
  selector.resetAreaOffset();
  verticalScrollbarSet.call(this);
  editorSetOffset.call(this);
}

function colResizerFinished(cRect, distance) {
  const { ci } = cRect;
  const { table, selector, data } = this;
  const { sci, eci } = selector.range;
  if (ci >= sci && ci <= eci) {
    for (let col = sci; col <= eci; col += 1) {
      data.cols.setWidth(col, distance);
    }
  } else {
    data.cols.setWidth(ci, distance);
  }

  table.render();
  selector.resetAreaOffset();
  horizontalScrollbarSet.call(this);
  editorSetOffset.call(this);
}

function dataSetCellText(text, state = "finished") {
  const {
    data,
    editor,
    contextMenu,
    autoFilter,
    sortFilter,
  } = this;
  const { ri, ci } = data.selector;
  if (!data.selector || ri === undefined || ci === undefined) {
    // console.warn('dataSetCellText: data.selector, ri 或 ci 无效', data.selector, ri, ci);
    return;
  }
  if (ri === -1 || ci === -1) return;
  const trigger =
    this.options.formula && this.options.formula.trigger
      ? this.options.formula.trigger
      : "$";
  // const cell = data.getCell(ri, ci);

  const inputText = editor.inputText;
  const trimmedText = text?.trim?.();
  if (editor.formulaCell && state === "finished") {
    const { ri, ci } = editor.formulaCell;
    data.setFormulaCellText(inputText, ri, ci, state);
    this.trigger("cell-edited", inputText, ri, ci);
    this.trigger("cell-edit-finished", text, ri, ci);
    editor.setFormulaCell(null);
    //The below condition is ro inject variable inside formula when there is only one formula and above case will handle for = sign, if there is more variable or ant this ese then it will go to text
  } else if (
    state === "finished" &&
    trimmedText?.startsWith(trigger) &&
    trimmedText?.split(" ").length === 1
  ) {
    const { ri, ci } = data.selector;
    data.setFormulaCellText(inputText, ri, ci, state);
    this.trigger("cell-edited", inputText, ri, ci);
    this.trigger("cell-edit-finished", text, ri, ci);
  } else if (!editor.formulaCell) {
    data.setFormulaCellText(text, ri, ci, state);
    this.trigger("cell-edited", text, ri, ci);
    this.trigger("cell-edit-finished", text, ri, ci);
  }
  if (autoFilter && ri !== null && ri !== undefined && ci !== null && ci !== undefined && autoFilter.includes(ri, ci)) {
    autoFilter.reload();
    sortFilter.reload();
  }
  contextMenu.hide();
}

function insertDeleteRowColumn(type) {
  const { data } = this;
  if (data.settings.mode === "read" || data.settings.mode === "preview" || data.settings.mode === "enabled") return;
  
  if (type === "insert-row") {
    data.insert("row");
  } else if (type === "delete-row") {
    data.delete("row");
  } else if (type === "insert-column") {
    data.insert("column");
  } else if (type === "delete-column") {
    data.delete("column");
  } else if (type === "delete-cell") {
    data.deleteCell();
  } else if (type === "delete-cell-format") {
    data.deleteCell("format");
  } else if (type === "delete-cell-text") {
    data.deleteCell("text");
  } else if (type === "cell-printable") {
    data.setSelectedCellAttr("printable", true);
  } else if (type === "cell-non-printable") {
    data.setSelectedCellAttr("printable", false);
  } else if (type === "cell-editable") {
    data.setSelectedCellAttr("editable", true);
  } else if (type === "cell-non-editable") {
    data.setSelectedCellAttr("editable", false);
  } else if (type === "set-data-cell") {
    data.setSelectedCellAttr("isDataCell", true);
  } else if (type === "cancel-data-cell") {
    data.setSelectedCellAttr("isDataCell", false);
  } else if (type === "cell-editable") {
    data.setSelectedCellAttr("editable", true);
    this.table.render();
  } else if (type === "cell-non-editable") {
    data.setSelectedCellAttr("editable", false);
    this.table.render();
  } else if (type === "cell-type-text") {
    data.setSelectedCellAttr("cellType", "none");
    this.table.render();
  } else if (type === "cell-type-date") {
    data.setSelectedCellAttr("cellType", "date");
    this.table.render();
  } else if (type === "cell-type-tree") {
    data.setSelectedCellAttr("cellType", "tree");
    this.table.render();
  } else if (type === "cell-type-popup") {
    data.setSelectedCellAttr("cellType", "popup");
    this.table.render();
  }
  clearClipboard.call(this);
  sheetReset.call(this);
}

function toolbarChange(type, value) {
  const { data, options } = this;
  const { cellConfigButtons = [] } = options;
  if (type === "undo") {
    this.undo();
  } else if (type === "redo") {
    this.redo();
  } else if (type === "print") {
    this.print.preview();
  } else if (type === "paintformat") {
    if (value === true) copy.call(this);
    else clearClipboard.call(this);
  } else if (type === "clearformat") {
    insertDeleteRowColumn.call(this, "delete-cell-format");
  } else if (type === "link") {
    // link
  } else if (type === "chart") {
    // chart
  } else if (type === "autofilter") {
    // filter
    autofilter.call(this);
  } else if (type === "freeze") {
    if (value) {
      const { ri, ci } = data.selector;
      this.freeze(ri, ci);
    } else {
      this.freeze(0, 0);
    }
  } else if (cellConfigButtons?.find((config) => config.tag === type)) {
    data.setSelectedCellAttr(type, value);
    sheetReset.call(this);
  } else if (type === "import") {
    this.data?.rootContext?.importWorkbook(value);
  } else {
    data.setSelectedCellAttr(type, value);
    if (type === "formula" && !data.selector.multiple()) {
      editorSet.call(this);
    }
    sheetReset.call(this);
  }
}

function sortFilterChange(ci, order, operator, value) {

  this.data.setAutoFilter(ci, order, operator, value);
  sheetReset.call(this);
}

function navigateToCell() {
  const cell = this.data.getSelectedCell();
  const { f } = cell ?? {};
  if (f) {
    const sheetToCellRef = f.match(SHEET_TO_CELL_REF_REGEX);
    if (sheetToCellRef?.length) {
      const [linkSheetName, cellRef] = sheetToCellRef[0]
        .replaceAll("'", "")
        .split("!");
      const [x, y] = expr2xy(cellRef);
      const sheetNames = this.data?.rootContext?.bottombar?.dataNames ?? [];
      const sheetIndex = sheetNames.indexOf(linkSheetName);
      if (sheetIndex !== -1) {
        this.data.rootContext.selectSheet(sheetIndex);
        this.selectCellAndFocus.call(this, y, x);
      }
    } else {
      const cellRef = f.match(CELL_REF_REGEX);
      if (cellRef?.length) {
        const [x, y] = expr2xy(cellRef[0]);
        this.selectCellAndFocus.call(this, y, x);
      }
    }
  }
}

function sheetInitEvents() {
  const {
    selector,
    overlayerEl,
    rowResizer,
    colResizer,
    verticalScrollbar,
    horizontalScrollbar,
    editor,
    contextMenu,
    toolbar,
    modalValidation,
    sortFilter,
    comment,
  } = this;
  // overlayer
  overlayerEl
    .on("mousemove", (evt) => {
      overlayerMousemove.call(this, evt);
    })
    .on("mousedown", (evt) => {
      if (!editor.formulaCell) {
        editor.clear();
      }
      contextMenu.hide();
      // the left mouse button: mousedown → mouseup → click
      // the right mouse button: mousedown → contenxtmenu → mouseup
      if (evt.buttons === 2) {
        if (this.data.xyInSelectedRect(evt.offsetX, evt.offsetY)) {
          contextMenu.show(evt);
        } else {
          overlayerMousedown.call(this, evt);
          contextMenu.show(evt);
        }
        evt.stopPropagation();
      } else if (evt.detail === 2) {
        this.dblclickHandler(evt);
      } else {
        overlayerMousedown.call(this, evt);
      }
    })
    .on("mousewheel.stop", (evt) => {
      const { className } = evt.target;
      if (className === `${cssPrefix}-overlayer`)
        overlayerMousescroll.call(this, evt);
    })
    .on("mouseout", (evt) => {
      const { offsetX, offsetY } = evt;
      if (offsetY <= 0) colResizer.hide();
      if (offsetX <= 0) rowResizer.hide();
    });

  selector.inputChange = (v) => {

    dataSetCellText.call(this, v, "input");
    editorSet.call(this);
  };

  // slide on mobile
  bindTouch(overlayerEl.el, {
    move: (direction, d) => {
      overlayerTouch.call(this, direction, d);
    },
  });

  // toolbar change
  toolbar.change = (type, value) => toolbarChange.call(this, type, value);

  // sort filter ok
  sortFilter.ok = (ci, order, o, v) =>
    sortFilterChange.call(this, ci, order, o, v);

  // resizer finished callback
  rowResizer.finishedFn = (cRect, distance) => {
    rowResizerFinished.call(this, cRect, distance);
  };
  colResizer.finishedFn = (cRect, distance) => {
    colResizerFinished.call(this, cRect, distance);
  };
  // resizer unhide callback
  rowResizer.unhideFn = (index) => {
    unhideRowsOrCols.call(this, "row", index);
  };
  colResizer.unhideFn = (index) => {
    unhideRowsOrCols.call(this, "col", index);
  };
  // scrollbar move callback
  verticalScrollbar.moveFn = (distance, evt) => {
    verticalScrollbarMove.call(this, distance, evt);
  };
  horizontalScrollbar.moveFn = (distance, evt) => {
    horizontalScrollbarMove.call(this, distance, evt);
  };
  // editor
  editor.change = (state, itext) => {
    if (itext?.trim?.()?.startsWith("=")) {
      const { ri, ci } = this.data.selector;
      if (!editor.formulaCell) {
        editor.setFormulaCell({ ri, ci });
      }
    } else {
      editor.setFormulaCell(null);
    }
    dataSetCellText.call(this, itext, state);
  };
  // modal validation
  modalValidation.change = (action, ...args) => {
    if (action === "save") {
      this.data.addValidation(...args);
    } else {
      this.data.removeValidation();
    }
  };
  // contextmenu
  contextMenu.itemClick = (type) => {
    const extendedContextMenus = this.options.extendedContextMenu;
    if (extendedContextMenus?.length) {
      const flattenedMenu = extendedContextMenus.reduce((acc, item) => {
        acc.push({ key: item.key, title: item.title, callback: item.callback });
        if (item.subMenus) {
          acc = acc.concat(item.subMenus);
        }
        return acc;
      }, []);
      const match = flattenedMenu?.find((menu) => menu.key === type);
      if (match) {
        const { ri, ci, range } = this.data.selector;
        const cell = this.data.getSelectedCell();
        match.callback?.call?.(this, ri, ci, range, cell, match);
      }
    }
    if (type === "validation") {
      modalValidation.setValue(this.data.getSelectedValidation());
    } else if (type === "copy") {
      copy.call(this);
    } else if (type === "cut") {
      cut.call(this);
    } else if (type === "paste") {
      paste.call(this, "all");
    } else if (type === "paste-value") {
      paste.call(this, "text");
    } else if (type === "paste-format") {
      paste.call(this, "format");
    } else if (type === "hide") {
      hideRowsOrCols.call(this);
    } else if (type === "add-comment" || type === "show-comment") {
      comment.show();
      contextMenu.hide();
    } else if (type === "navigate") {
      navigateToCell.call(this);
    } else {
      insertDeleteRowColumn.call(this, type);
    }
  };

  bind(window, "resize", () => {
    this.reload();
  });

  bind(window, "click", (evt) => {
    this.focusing = overlayerEl.contains(evt.target);
  });

  bind(window, "paste", (evt) => {
    if (!this.focusing) return;
    paste.call(this, "all", evt);
    evt.preventDefault();
  });

  bind(window, "copy", (evt) => {
    if (!this.focusing) return;
    copy.call(this, evt);
    evt.preventDefault();
  });

  // for selector
  bind(window, "keydown", (evt) => {
    if (!this.focusing) return;
    const keyCode = evt.keyCode || evt.which;
    const { key, ctrlKey, shiftKey, metaKey } = evt;

    if (ctrlKey || metaKey) {
      // const { sIndexes, eIndexes } = selector;
      // let what = 'all';
      // if (shiftKey) what = 'text';
      // if (altKey) what = 'format';
      switch (keyCode) {
        case 90:
          // undo: ctrl + z
          this.undo();
          evt.preventDefault();
          break;
        case 89:
          // redo: ctrl + y
          this.redo();
          evt.preventDefault();
          break;
        case 67:
          // ctrl + c
          // => copy
          // copy.call(this);
          // evt.preventDefault();
          break;
        case 88:
          // ctrl + x
          cut.call(this);
          evt.preventDefault();
          break;
        case 85:
          // ctrl + u
          toolbar.trigger("underline");
          evt.preventDefault();
          break;
        case 86:
          // ctrl + v
          // => paste
          // evt.preventDefault();
          break;
        case 37:
          // ctrl + left
          selectorMove.call(this, shiftKey, "row-first");
          evt.preventDefault();
          break;
        case 38:
          // ctrl + up
          selectorMove.call(this, shiftKey, "col-first");
          evt.preventDefault();
          break;
        case 39:
          // ctrl + right
          selectorMove.call(this, shiftKey, "row-last");
          evt.preventDefault();
          break;
        case 40:
          // ctrl + down
          selectorMove.call(this, shiftKey, "col-last");
          evt.preventDefault();
          break;
        case 32:
          // ctrl + space, all cells in col
          selectorSet.call(this, false, -1, this.data.selector.ci, false);
          evt.preventDefault();
          break;
        case 66:
          // ctrl + B
          toolbar.trigger("bold");
          break;
        case 73:
          // ctrl + I
          toolbar.trigger("italic");
          break;
        case 65:
          // ctrl + A
          selector.set(-1, -1, true);
          break;
        case 219:
          // Ctrl + [
          navigateToCell.call(this);
        default:
          break;
      }
    } else {
      switch (keyCode) {
        case 32:
          if (shiftKey) {
            // shift + space, all cells in row
            selectorSet.call(this, false, this.data.selector.ri, -1, false);
          }
          break;
        case 27: // esc
          contextMenu.hide();
          clearClipboard.call(this);
          break;
        case 37: // left
          selectorMove.call(this, shiftKey, "left");
          evt.preventDefault();
          break;
        case 38: // up
          selectorMove.call(this, shiftKey, "up");
          evt.preventDefault();
          break;
        case 39: // right
          selectorMove.call(this, shiftKey, "right");
          evt.preventDefault();
          break;
        case 40: // down
          selectorMove.call(this, shiftKey, "down");
          evt.preventDefault();
          break;
        case 9: // tab
          editor.clear();
          // shift + tab => move left
          // tab => move right
          selectorMove.call(this, false, shiftKey ? "left" : "right");
          evt.preventDefault();
          break;
        case 13: // enter
          editor.clear();
          // shift + enter => move up
          // enter => move down
          selectorMove.call(this, false, shiftKey ? "up" : "down");
          evt.preventDefault();
          break;
        case 8: // backspace
          insertDeleteRowColumn.call(this, "delete-cell-text");
          evt.preventDefault();
          break;
        default:
          break;
      }

      if (key === "Delete") {
        insertDeleteRowColumn.call(this, "delete-cell-text");
        evt.preventDefault();
      } else if (
        (keyCode >= 65 && keyCode <= 90) ||
        (keyCode >= 48 && keyCode <= 57) ||
        (keyCode >= 96 && keyCode <= 105) ||
        evt.key === "="
      ) {
        // 在预览模式下检查单元格是否可编辑
        const { data } = this;
        const cell = data.getSelectedCell();
        const mode = data.settings?.mode;
        if ((mode === "preview" || mode === "normal" || mode === "read" || mode === "enabled") && cell && cell.editable === false) {
          return;
        }
        
        dataSetCellText.call(this, evt.key, "input");
        editorSet.call(this);
      } else if (keyCode === 113) {
        // F2
        // 在预览模式下检查单元格是否可编辑
        const { data } = this;
        const cell = data.getSelectedCell();
        const mode = data.settings?.mode;
        if (mode === "preview" || mode === "normal" || mode === "read" || mode === "enabled") {
          if (cell && cell.editable === false) {
            return;
          }
        }
        editorSet.call(this);
      }
    }
  });

  bind(window, "keyup", (evt) => {
    const keyCode = evt.keyCode || evt.which;
    if (keyCode === 16 || keyCode === 17) {
      const range = this.selector.range;
      if (range.sri === range.eri && range.sci === range.eci) return;
      else {
        const cell = this.data.getCell(range.sri, range.sci);
        this.trigger("cells-selected", cell, range);
      }
    }
  });
  bind(document, "visibilitychange", (evt) => {
    if (document.visibilityState === "hidden") {
      this.data?.clipboard?.clear();
    }
  });

  this.on("context-menu-action", (params) => {
    const { action: [type], range: { sri: ri, sci: ci } } = params;
    if (/^cell-type-\w+/.test(type)) {
      if (ri === undefined || ci === undefined || ri < 0 || ci < 0) return;
      const cell = this.data.rows.getCell(ri, ci);
      if (cell.isDataListRiHeader || ri === this.data.rows.dataListRange.sri) {
        const cellType = type.split("-")[2]
        this.data.rows.syncDataListColumnType(ci, cellType);
      }
    }
  });
}

export default class Sheet {
  constructor(targetEl, data, options = {}) {
    this.options = options;
    this.sheetIndex = data.settings?.sheetIndex || 1;
    this.data = data;
    this.eventMap = createEventEmitter();
    const { view, showToolbar, showContextmenu } = data.settings;
    this.el = h("div", `${cssPrefix}-sheet`);
    this.toolbar = new Toolbar(this, data, view.width, !showToolbar);
    this.print = new Print(data);
    targetEl.children(this.toolbar.el, this.el, this.print.el);
    // table
    this.tableEl = h("canvas", `${cssPrefix}-table`);
    // resizer
    this.rowResizer = new Resizer(false, data.rows.minHeight);
    this.colResizer = new Resizer(true, data.cols.minWidth);
    // scrollbar
    this.verticalScrollbar = new Scrollbar(true);
    this.horizontalScrollbar = new Scrollbar(false);
    // editor
    this.editor = new Editor(
      this.data,
      formulas,
      () => this.getTableOffset(),
      data.rows.height,
      this.options
    );
    // data validation
    this.modalValidation = new ModalValidation();
    // contextMenu
    this.contextMenu = new ContextMenu(
      this,
      () => this.getRect(),
      !showContextmenu,
      options.extendedContextMenu
    );
    // selector
    this.selector = new Selector(data);
    this.overlayerCEl = h("div", `${cssPrefix}-overlayer-content`).children(
      this.editor.el,
      this.selector.el
    );
    this.overlayerEl = h("div", `${cssPrefix}-overlayer`).child(
      this.overlayerCEl
    );
    // sortFilter
    this.sortFilter = new SortFilter();
    this.comment = new Comment(this, () => this.getRect());
    this.hoverTimer = null;
    
    // 导入自定义组件
    const TreeSelector = options.components?.TreeSelector || this.options.TreeSelector;
    const PopupDialog = options.components?.PopupDialog || this.options.PopupDialog;
    
    // 导入或创建树形选择器
    this.treeSelector = null;
    if (TreeSelector) {
      try {
        this.treeSelector = new TreeSelector(options.treeSelector || {});
        // 如果创建了树形选择器，将它添加到DOM中
        if (this.treeSelector) {
          this.el.child(this.treeSelector.el);
        }
      } catch (e) {
        console.error('创建树形选择器失败:', e);
      }
    } else {
      // 如果没有提供TreeSelector，使用内置的
      try {
        const TreeSelectorInternal = require('./tree-selector').default;
        this.treeSelector = new TreeSelectorInternal(options.treeSelector || {});
        this.el.child(this.treeSelector.el);
      } catch (e) {
        console.error('创建内置树形选择器失败:', e);
      }
    }
    
    // 导入或创建弹窗
    this.popupDialog = null;
    if (PopupDialog) {
      try {
        this.popupDialog = new PopupDialog(options.popupDialog || {});
        // 弹窗直接添加到body，不需要添加到el中
      } catch (e) {
        console.error('创建弹窗失败:', e);
      }
    } else {
      // 如果没有提供PopupDialog，使用内置的
      try {
        const PopupDialogInternal = require('./popup-dialog').default;
        this.popupDialog = new PopupDialogInternal(options.popupDialog || {});
      } catch (e) {
        console.error('创建内置弹窗失败:', e);
      }
    }
    
    // 创建树形编辑器
    try {
      const TreeEditor = require('./tree-editor').default;
      this.treeEditor = new TreeEditor(options.treeEditor || {});
      this.treeEditor.change = (params) => {
        const { ri, ci, treeData, placeholder } = params;
        // 更新单元格的treeData
        const cell = this.data.getCell(ri, ci) || {};
        cell.treeData = treeData;
        // 更新单元格文本为提示词
        cell.text = placeholder || '点击选择';
        // 确保单元格是树形类型
        cell.cellType = 'tree';
        // 更新到数据层
        this.data.setCellData(ri, ci, cell);
        // 重新渲染表格
        this.table.render();
      };
      document.body.appendChild(this.treeEditor.el.el);
    } catch (e) {
      console.error('创建树形编辑器失败:', e);
    }
    
    // 创建弹窗内容编辑器
    try {
      const PopupEditor = require('./popup-editor').default;
      this.popupEditor = new PopupEditor(options.popupEditor || {});
      this.popupEditor.change = (params) => {
        const { ri, ci, popupData } = params;
        // 更新单元格的popupData
        const cell = this.data.getCell(ri, ci) || {};
        cell.popupData = popupData;
        // 更新单元格文本
        cell.text = '点击打开';
        // 确保单元格是弹窗类型
        cell.cellType = 'popup';
        // 更新到数据层
        this.data.setCellData(ri, ci, cell);
        // 重新渲染表格
        this.table.render();
      };
      document.body.appendChild(this.popupEditor.el.el);
    } catch (e) {
      console.error('创建弹窗编辑器失败:', e);
    }
    
    // 创建日期选择器
    this.datepicker = new Datepicker();
    document.body.appendChild(this.datepicker.el.el);
    
    // root element
    this.el.children(
      this.tableEl,
      this.overlayerEl.el,
      this.rowResizer.el,
      this.colResizer.el,
      this.verticalScrollbar.el,
      this.horizontalScrollbar.el,
      this.contextMenu.el,
      this.modalValidation.el,
      this.sortFilter.el,
      this.comment.el
    );
    // table
    this.table = new Table(this.tableEl.el, data, this.options);
    sheetInitEvents.call(this);
    sheetReset.call(this);
    // init selector [0, 0]
    selectorSet.call(this, false, 0, 0);
    if (this.options.mode === "read") this.selector.hide();
    
    // 在预览模式下隐藏工具栏
    if ((options.mode === 'preview' || options.mode === 'enabled') && this.toolbar) {
      this.toolbar.el.hide();
    }
    
    // 添加事件监听器
    this.on('show-tree-selector', ({ ri, ci, cell }) => {
      if (this.treeSelector) {
        // 设置树形数据
        this.treeSelector.setItems(cell.treeData || []);
        
        // 设置回调函数
        this.treeSelector.setChange((selectedItem) => {
          this.data.setCellText(ri, ci, selectedItem.label);
          this.table.render();
        });
        
        // 显示树形选择器
        this.treeSelector.show();
      }
    });
    
    this.on('show-tree-editor', ({ ri, ci, cell }) => {
      if (this.treeEditor) {
        // 设置当前编辑的单元格
        this.treeEditor.setCell(ri, ci, cell);
        
        // 计算位置 - 尽量居中显示
        const viewRect = this.getRect();
        const editorWidth = this.treeEditor.width;
        const editorHeight = this.treeEditor.height;
        
        // 默认居中
        let left = viewRect.left + (viewRect.width - editorWidth) / 2;
        let top = viewRect.top + (viewRect.height - editorHeight) / 2;
        
        // 确保不超出视窗
        left = Math.max(0, Math.min(left, window.innerWidth - editorWidth));
        top = Math.max(0, Math.min(top, window.innerHeight - editorHeight));
        
        // 设置编辑器位置
        this.treeEditor.el
          .css('position', 'absolute')
          .css('left', `${left}px`)
          .css('top', `${top}px`);
        
        // 显示编辑器
        this.treeEditor.show();
      }
    });
    
    this.on('show-popup', ({ ri, ci, cell }) => {
      if (this.popupDialog) {
        // 获取弹窗数据
        const popupData = cell.popupData || options.popupData || {
          title: '弹窗标题',
          width: 400,
          height: 300,
          content: '<div style="padding:10px">弹窗内容</div>'
        };
        
        // 设置标题
        this.popupDialog.setTitle(popupData.title);
        
        // 设置尺寸
        this.popupDialog.setSize(popupData.width, popupData.height);
        
        // 设置内容
        this.popupDialog.setContent(popupData.content);
        
        // 设置确认回调
        this.popupDialog.onConfirm = () => {
          // 更新单元格的选择状态
          this.data.setCellProperty(ri, ci, 'selected', true);
          // 更新单元格文本
          this.data.setCellText(ri, ci, '已选择');
          // 重新渲染表格
          this.table.render();
        };
        
        // 显示弹窗
        this.popupDialog.show();
      }
    });
    
    this.on('show-popup-editor', ({ ri, ci, cell }) => {
      if (this.popupEditor) {
        // 设置当前编辑的单元格
        this.popupEditor.setCell(ri, ci, cell);
        
        // 计算位置 - 尽量居中显示
        const viewRect = this.getRect();
        const editorWidth = this.popupEditor.width;
        const editorHeight = this.popupEditor.height;
        
        // 默认居中
        let left = viewRect.left + (viewRect.width - editorWidth) / 2;
        let top = viewRect.top + (viewRect.height - editorHeight) / 2;
        
        // 确保不超出视窗
        left = Math.max(0, Math.min(left, window.innerWidth - editorWidth));
        top = Math.max(0, Math.min(top, window.innerHeight - editorHeight));
        
        // 设置编辑器位置
        this.popupEditor.el
          .css('position', 'absolute')
          .css('left', `${left}px`)
          .css('top', `${top}px`);
        
        // 显示弹窗编辑器
        this.popupEditor.show();
      }
    });
    
    this.on('show-datepicker', ({ ri, ci, cell }) => {
      if (this.datepicker) {
        // 设置当前编辑的单元格
        this.datepicker.setCell(ri, ci, cell);
        
        // 设置回调函数
        this.datepicker.setChange((formattedDate) => {
          this.data.setCellText(ri, ci, formattedDate);
          this.table.render();
        });
        
        // 显示弹窗
        this.datepicker.show();
      }
    });
  }

  on(eventName, func) {
    this.eventMap.on(eventName, func);
    return this;
  }

  trigger(eventName, ...args) {
    const { eventMap } = this;
    eventMap.fire(eventName, args);
  }

  resetData(data) {
    // before
    this.editor.clear();
    // after
    this.data = data;
    verticalScrollbarSet.call(this);
    horizontalScrollbarSet.call(this);
    this.toolbar.resetData(data);
    this.print.resetData(data);
    this.selector.resetData(data);
    this.table.resetData(data);
    this.editor.resetData(data);
    setTimeout(() => {
      this.trigger("grid-load", data);
    });
  }

  loadData(data) {
    this.data.setData(data);
    sheetReset.call(this);
    return this;
  }

  // freeze rows or cols
  freeze(ri, ci) {
    const { data } = this;
    data.setFreeze(ri, ci);
    sheetReset.call(this);
    return this;
  }

  undo() {
    this.data.undo();
    sheetReset.call(this);
  }

  redo() {
    this.data.redo();
    sheetReset.call(this);
  }

  reload() {
    sheetReset.call(this);
    return this;
  }

  getRect() {
    const { data } = this;
    return { width: data.viewWidth(), height: data.viewHeight() };
  }

  getTableOffset() {
    const { rows, cols } = this.data;
    const { width, height } = this.getRect();
    return {
      width: width - cols.indexWidth,
      height: height - rows.height,
      left: cols.indexWidth,
      top: rows.height,
    };
  }

  selectCell(ri, ci) {
    selectorSet.call(this, false, ri, ci);
  }

  selectCellAndFocus(ri, ci) {
    selectorSet.call(this, false, ri, ci);
    scrollbarMove.call(this);
  }

  selectCellsAndFocus(range) {
    const { sri, eri, sci, eci } = range ?? {};
    if (
      isNaN(parseInt(sri)) ||
      isNaN(parseInt(eri)) ||
      isNaN(parseInt(sci)) ||
      isNaN(parseInt(eci))
    ) {
      return;
    }
    if (sri === eri && sci === eci) selectorSet.call(this, false, sri, sci);
    else {
      selectorSet.call(this, false, sri, sci);
      for (let ri = sri; ri <= eri; ri++) {
        for (let ci = sci; ci <= eci; ci++) {
          selectorSet.call(this, true, ri, ci, true, true);
        }
      }
      selectorSet.call(this, true, eri, eci, true, false);
    }
    scrollbarMove.call(this);
  }

  dblclickHandler(evt) {
    const { data } = this;
    const { ri, ci } = data.getCellRectByXY(evt.offsetX, evt.offsetY);
    const cell = this.data.rows.getCell(ri, ci);
    // 检查当前模式
    const mode = this.data.settings?.mode;
    // 如果单元格不可编辑且不是设计模式
    if (cell && cell.editable === false && mode !== 'design') {
      return;
    }
    // 根据单元格类型处理不同的编辑方式
    if (cell && cell.cellType && cell.cellType !== "none" && cell.cellType !== "text") {
      // 选择单元格以确保位置正确
      this.selectCell(ri, ci);
      // 计算单元格位置信息
      const rect = data.cellRect(ri, ci);
      if (!rect) {
        editorSet.call(this);
        return;
      }
      // 根据不同类型使用不同的编辑器或弹窗
      if (cell.cellType === 'date') {
        // 日期选择器处理逻辑
        this.trigger('show-datepicker', { ri, ci, cell });
        return;
      } else if (cell.cellType === 'tree') {
        // 触发树形选择器事件
        if (mode === 'design') {
          // 设计模式：允许编辑树形数据
          this.trigger('show-tree-editor', { ri, ci, cell });
        } else {
          // 预览模式：只允许选择
          this.trigger('show-tree-selector', { ri, ci, cell });
        }
        return; // 不要继续使用默认编辑器
      } else if (cell.cellType === 'popup') {
        // 触发弹窗事件
        if (mode === 'design') {
          // 设计模式：允许编辑弹窗内容
          this.trigger('show-popup-editor', { ri, ci, cell });
        } else {
          // 预览模式：只显示弹窗
          this.trigger('show-popup', { ri, ci, cell });
        }
        return; // 不要继续使用默认编辑器
      }
    }
    // 常规编辑处理
    editorSet.call(this);
  }
  
}
