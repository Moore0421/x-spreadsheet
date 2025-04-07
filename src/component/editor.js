//* global window */
import { h } from "./element";
import Suggest from "./suggest";
import Datepicker from "./datepicker";
import { cssPrefix } from "../config";
import MentionMenu from "./mentionmenu";
import { DEFAULT_ROW_HEIGHT } from "../constants";

function resetTextareaSize() {
  const { inputText } = this;
  if (!/^\s*$/.test(inputText)) {
    const { textlineEl, textEl, areaOffset } = this;
    const txts = inputText?.toString?.()?.split?.("\n");
    const maxTxtSize = Math.max(...txts.map((it) => it.length));
    const tlOffset = textlineEl.offset();
    const fontWidth = tlOffset.width / inputText.length;
    const tlineWidth = (maxTxtSize + 1) * fontWidth + 5;
    const maxWidth = this.viewFn().width - areaOffset.left - fontWidth;
    let h1 = txts.length;
    if (tlineWidth > areaOffset.width) {
      let twidth = tlineWidth;
      if (tlineWidth > maxWidth) {
        twidth = maxWidth;
        h1 += parseInt(tlineWidth / maxWidth, 10);
        h1 += tlineWidth % maxWidth > 0 ? 1 : 0;
      }
      textEl.css("width", `${twidth}px`);
    }
    h1 *= this.rowHeight;
    if (h1 > areaOffset.height) {
      textEl.css("height", `${h1}px`);
    }
  }
}

function insertText({ target }, itxt) {
  const { value, selectionEnd } = target;
  const ntxt = `${value.slice(0, selectionEnd)}${itxt}${value.slice(selectionEnd)}`;
  target.value = ntxt;
  target.setSelectionRange(selectionEnd + 1, selectionEnd + 1);

  this.inputText = ntxt;
  this.textlineEl.html(ntxt);
  resetTextareaSize.call(this);
}

function keydownEventHandler(evt) {
  const { keyCode, altKey } = evt;
  if (keyCode !== 13 && keyCode !== 9) evt.stopPropagation();
  if (keyCode === 13 && altKey) {
    insertText.call(this, evt, "\n");
    const { ri } = this.data?.selector ?? {};
    this.data.setRowHeight(ri, this.data.getRowHeight(ri) + DEFAULT_ROW_HEIGHT);
    evt.stopPropagation();
  }
  if (keyCode === 13 && !altKey) evt.preventDefault();
}

function mentionMenuSearch(text) {
  const { textEl, trigger, mention } = this;
  const caretPos = textEl.el.selectionStart;
  let start = caretPos - 1;
  while (start >= 0 && /\S/.test(text[start])) {
    start--;
  }
  start++;

  let end = caretPos;
  while (end < text.length && /\S/.test(text[end])) {
    end++;
  }

  let currentWord = text.slice(start, end);
  if (currentWord?.startsWith(trigger)) {
    mention.search(currentWord);
  } else {
    mention.hide();
  }
}

function inputEventHandler(evt) {
  const v = evt.target.value;
  const { cell, ri, ci, data } = this;
  
  // 如果单元格不可编辑且不是设计模式，则阻止输入
  if (cell && cell.editable === false && data.settings?.mode !== 'design') {
    return;
  }
  
  // 其他处理逻辑保持不变
  this.inputText = v;
  if (evt.type === 'input') {
    this.change("input", v);
  }
  
  const { suggest, textlineEl, validator, mention, trigger } = this;
  if (cell !== null) {
    if (validator) {
      // 如果是下拉列表类型
      if (validator.type === 'list') {
        // 只允许从下拉列表中选择,不允许直接输入
        suggest.search(v);
        return;
      }
      // 验证输入的内容
      const [isValid] = validator.validate(v);
      if (!isValid) {
        // 输入无效时,恢复原来的值
        evt.target.value = this.inputText; 
        return;
      }
    }
  }
  if (cell !== null) {
    if (
      ("editable" in cell && cell.editable === true) ||
      cell.editable === undefined ||
      data.settings.mode === 'design'
    ) {
      if (validator) {
        if (validator.type === "list") {
          suggest.search(v);
        } else {
          suggest.hide();
        }
      } else {
        const start = v.lastIndexOf("=");
        if (start !== -1) {
          suggest.search(v.substring(start + 1));
        } else {
          suggest.hide();
        }
      }
      if (trigger) {
        if (v?.includes(trigger)) {
          suggest.hide();
          mentionMenuSearch.call(this, v);
        } else {
          mention.hide();
        }
      }
      textlineEl.html(v);
      resetTextareaSize.call(this);
    } else {
      evt.target.value = cell.text ?? "";
    }
  } else {
    if (validator) {
      if (validator.type === "list") {
        suggest.search(v);
      } else {
        suggest.hide();
      }
    } else {
      const start = v.lastIndexOf("=");
      if (start !== -1) {
        suggest.search(v.substring(start + 1));
      } else {
        suggest.hide();
      }
      if (trigger) {
        if (v?.includes(trigger)) {
          suggest.hide();
          mentionMenuSearch.call(this, v);
        } else {
          mention.hide();
        }
      }
    }
    textlineEl.html(v);
    resetTextareaSize.call(this);
  }
}

function setTextareaRange(position) {
  const { el } = this.textEl;
  setTimeout(() => {
    el.focus();
    el.setSelectionRange(position, position);
  }, 0);
}

function setText(text, position) {
  const { textEl, textlineEl } = this;
  // firefox bug
  textEl.el.blur();

  textEl.val(text);
  textlineEl.html(text);
  setTextareaRange.call(this, position);
}

function suggestItemClick(it) {
  const { inputText, validator } = this;
  let position = 0;
  if (validator && validator.type === "list") {
    this.inputText = it;
    position = this.inputText.length;
  } else {
    const start = inputText.lastIndexOf("=");
    const sit = inputText.substring(0, start + 1);
    let eit = inputText.substring(start + 1);
    if (eit.indexOf(")") !== -1) {
      eit = eit.substring(eit.indexOf(")"));
    } else {
      eit = "";
    }
    this.inputText = `${sit + it.key}(`;
    // console.log('inputText:', this.inputText);
    position = this.inputText.length;
    this.inputText += `)${eit}`;
  }
  setText.call(this, this.inputText, position);
}

function resetSuggestItems() {
  this.suggest.setItems(this.formulas);
}

function dateFormat(d) {
  let month = d.getMonth() + 1;
  let date = d.getDate();
  if (month < 10) month = `0${month}`;
  if (date < 10) date = `0${date}`;
  return `${d.getFullYear()}-${month}-${date}`;
}

function mentionInputHandler(item) {
  const { value } = item;
  const { textEl } = this;
  const caretPos = textEl.el.selectionStart;
  const text = this.inputText;
  let start = caretPos - 1;
  while (start >= 0 && /\S/.test(text[start])) {
    start--;
  }
  start++;

  let end = caretPos;
  while (end < text.length && /\S/.test(text[end])) {
    end++;
  }

  let currentWord = text.slice(start, end);

  let word = "";
  if (currentWord?.includes(this.trigger) && !currentWord?.includes(".")) {
    word = `#${value}`;
  } else if (
    currentWord?.includes(this.trigger) &&
    currentWord?.includes(".")
  ) {
    const index = currentWord.lastIndexOf(".");
    const lastText = currentWord.substring(0, index + 1);
    word = `${lastText}${value}`;
  }

  let newText =
    text.slice(0, start) + `${word?.replaceAll(" ", "_")}` + text.slice(end);

  this.inputText = newText;
  const position = start + word?.length ?? this.inputText?.length;
  setText.call(this, this.inputText, position);
  resetTextareaSize.call(this);
}

export default class Editor {
  constructor(data, formulas, viewFn, rowHeight, options = {}) {
    this.options = options;
    this.viewFn = viewFn;
    this.rowHeight = rowHeight;
    this.formulas = formulas;
    this.data = data;
    this.suggest = new Suggest(formulas, (it) => {
      suggestItemClick.call(this, it);
    });
    this.trigger = options.mentionProgress?.trigger;
    this.mention = new MentionMenu(options.mentionProgress, (item) => {
      mentionInputHandler.call(this, item);
    });
    this.datepicker = new Datepicker();
    this.datepicker.change((d) => {
      // console.log('d:', d);
      this.setText(dateFormat(d));
      this.clear();
    });
    this.areaEl = h("div", `${cssPrefix}-editor-area`)
      .children(
        (this.textEl = h("textarea", "")
          .on("input", (evt) => inputEventHandler.call(this, evt))
          .on("paste.stop", () => {})
          .on("keydown", (evt) => keydownEventHandler.call(this, evt))),
        (this.textlineEl = h("div", "textline")),
        this.suggest.el,
        this.datepicker.el,
        this.mention.el
      )
      .on("mousemove.stop", () => {})
      .on("mousedown.stop", () => {});
    this.el = h("div", `${cssPrefix}-editor`).child(this.areaEl).hide();
    this.suggest.bindInputEvents(this.textEl);
    this.mention.bindInputEvents(this.textEl);

    this.areaOffset = null;
    this.freeze = { w: 0, h: 0 };
    this.cell = null;
    this.inputText = "";
    this.change = () => {};
    this.formulaCell = null;
  }

  setFreezeLengths(width, height) {
    this.freeze.w = width;
    this.freeze.h = height;
  }

  clear() {
    // const { cell } = this;
    // const cellText = (cell && cell.text) || '';
    if (this.inputText !== "") {
      this.change("finished", this.inputText);
    }
    this.cell = null;
    this.areaOffset = null;
    this.inputText = "";
    this.el.hide();
    this.textEl.val("");
    this.textlineEl.html("");
    resetSuggestItems.call(this);
    this.datepicker.hide();
  }

  setOffset(offset, suggestPosition = "top") {
    const { textEl, areaEl, suggest, freeze, el, mention } = this;
    if (offset) {
      this.areaOffset = offset;
      const { left, top, width, height, l, t } = offset;
      // console.log('left:', left, ',top:', top, ', freeze:', freeze);
      const elOffset = { left: 0, top: 0 };
      // top left
      if (freeze.w > l && freeze.h > t) {
        //
      } else if (freeze.w < l && freeze.h < t) {
        elOffset.left = freeze.w;
        elOffset.top = freeze.h;
      } else if (freeze.w > l) {
        elOffset.top = freeze.h;
      } else if (freeze.h > t) {
        elOffset.left = freeze.w;
      }
      el.offset(elOffset);
      areaEl.offset({
        left: left - elOffset.left - 0.8,
        top: top - elOffset.top - 0.8,
      });
      textEl.offset({ width: width - 9 + 0.8, height: height - 3 + 0.8 });
      const sOffset = { left: 0 };
      sOffset[suggestPosition] = height;
      suggest.setOffset(sOffset);
      suggest.hide();
      mention.setOffset(sOffset);
      mention.hide();
    }
  }

  setCell(cell, validator) {
    // 检查是否在预览模式且单元格不可编辑
    const { data } = this;
    if ((data.settings.mode === 'preview' || data.settings.mode === 'normal' || data.settings.mode === 'read') && cell && cell.editable === false) {
      return;
    }

    const { el, textEl, textlineEl } = this;
    
    // 确保cell至少是一个空对象
    cell = cell || {};
    this.cell = cell;
    const { ri, ci } = data?.selector || {};
    this.ri = ri;
    this.ci = ci;
    
    const editValueFormatter = this.options.editValueFormatter;
    const { datepicker, suggest } = this;
    el.show();
    let text = "";
    if (editValueFormatter) {
      text = editValueFormatter({ ...this, cell }) ?? cell?.text ?? "";
    } else {
      text = cell?.f || (cell?.text ?? "");
    }
    this.setText(text);
    
    // 添加安全检查，确保validations存在且有getValidator方法
    if (data && data.validations && typeof data.validations.getValidator === 'function') {
      this.validator = validator;
    } else {
      this.validator = null; // 如果方法不存在，设置为null
    }
    
    // 在设计模式下，即使是不可编辑单元格也可以编辑
    if (data.settings.mode === 'design') {
      textEl.removeAttr("readonly");
    } else if (cell && cell.editable === false) {
      textEl.attr("readonly", "readonly");
    } else if (data && "editable" in data && data.editable === false) {
      textEl.attr("readonly", "readonly");
    } else {
      textEl.removeAttr("readonly");
    }
    
    if (cell?.text?.length === 1) {
      inputEventHandler.call(this, { target: { value: text } });
    }
  }

  setFormulaCell(cellRef) {
    this.formulaCell = cellRef;
  }

  setText(text) {
    this.inputText = text;
    setText.call(this, text, text.length);
    resetTextareaSize.call(this);
  }

  resetData(d) {
    this.data = d;
  }
}
