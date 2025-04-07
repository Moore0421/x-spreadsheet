/* global window */

import Align from "./align";
import Valign from "./valign";
import Autofilter from "./autofilter";
import Bold from "./bold";
import Grid from "./grid";
import Italic from "./italic";
import Strike from "./strike";
import Underline from "./underline";
import Border from "./border";
import Clearformat from "./clearformat";
import Paintformat from "./paintformat";
import TextColor from "./text_color";
import FillColor from "./fill_color";
import FontSize from "./font_size";
import Font from "./font";
import Format from "./format";
import Formula from "./formula";
import Freeze from "./freeze";
import Merge from "./merge";
import Redo from "./redo";
import Undo from "./undo";
import Print from "./print";
import Textwrap from "./textwrap";
import More from "./more";
import Item from "./item";

import { h } from "../element";
import { cssPrefix } from "../../config";
import { bind } from "../event";
import CellConfigButtons from "./cellConfigButtons";
import Import from "./import";
import { generateUniqueId } from "../../utils";

function buildDivider() {
  return h("div", `${cssPrefix}-toolbar-divider`);
}

function initBtns2() {
  this.btns2 = [];
  this.items.forEach((it) => {
    if (Array.isArray(it)) {
      it.forEach(({ el }) => {
        const rect = el.box();
        const { marginLeft, marginRight } = el.computedStyle();
        this.btns2.push([
          el,
          rect.width + parseInt(marginLeft, 10) + parseInt(marginRight, 10),
        ]);
      });
    } else {
      const rect = it.box();
      const { marginLeft, marginRight } = it.computedStyle();
      this.btns2.push([
        it,
        rect.width + parseInt(marginLeft, 10) + parseInt(marginRight, 10),
      ]);
    }
  });
}

function moreResize() {
  try {
    const { el, btns, moreEl, btns2 } = this;
    const { moreBtns, contentEl } = moreEl.dd;
    el.css("width", `${this.widthFn()}px`);
    const elBox = el.box();

    let sumWidth = 160;
    let sumWidth2 = 12;
    const list1 = [];
    const list2 = [];
    btns2.forEach(([it, w], index) => {
      sumWidth += w;
      if (index === btns2.length - 1 || sumWidth < elBox.width) {
        list1.push(it);
      } else {
        sumWidth2 += w;
        list2.push(it);
      }
    });
    btns.html("").children(...list1);
    moreBtns.html("").children(...list2);
    contentEl.css("width", `${sumWidth2}px`);
    if (list2.length > 0) {
      moreEl.show();
    } else {
      moreEl.hide();
    }
  } catch (e) {}
}

function genBtn(it) {
  const btn = new Item();
  btn.el.on("click", () => {
    const range = this.sheet.selector.range;
    this.sheet?.trigger("toolbar-action", {
      action: [it.tip],
      range,
    });
    if (it.onClick) it.onClick(this.data.getData(), this.data);
  });
  btn.tip = it.tip || "";

  let { el } = it;

  if (it.icon) {
    el = h("img").attr("src", it.icon);
  }

  if (el) {
    const icon = h("div", `${cssPrefix}-icon`);
    icon.child(el);
    btn.el.child(icon);
  }

  return btn;
}

export default class Toolbar {
  constructor(sheetContext = {}, data, widthFn, isHide = false) {
    this.data = data;
    this.change = () => {};
    this.widthFn = widthFn;
    this.isHide = isHide;
    this.sheet = sheetContext;
    const style = data.defaultStyle();
    this.items = [
      [
        {
          id: "import",
          btn: new Import(),
        },
      ],
      [
        { id: "undo", btn: (this.undoEl = new Undo()) },
        { id: "redo", btn: (this.redoEl = new Redo()) },
        { id: "print", btn: new Print() },
        { id: "paintFormat", btn: (this.paintformatEl = new Paintformat()) },
        { id: "clearFormat", btn: (this.clearformatEl = new Clearformat()) },
      ],
      [{ id: "format", btn: (this.formatEl = new Format()) }],
      [
        { id: "font", btn: (this.fontEl = new Font()) },
        { id: "fontSize", btn: (this.fontSizeEl = new FontSize()) },
      ],
      [
        { id: "bold", btn: (this.boldEl = new Bold()) },
        { id: "italic", btn: (this.italicEl = new Italic()) },
        { id: "underline", btn: (this.underlineEl = new Underline()) },
        { id: "strike", btn: (this.strikeEl = new Strike()) },
        {
          id: "textColor",
          btn: (this.textColorEl = new TextColor(style.color)),
        },
      ],
      [
        {
          id: "fillColor",
          btn: (this.fillColorEl = new FillColor(style.bgcolor)),
        },
        { id: "border", btn: (this.borderEl = new Border()) },
        { id: "merge", btn: (this.mergeEl = new Merge()) },
        { id: "grid", btn: (this.gridEl = new Grid()) },
      ],
      [
        { id: "align", btn: (this.alignEl = new Align(style.align)) },
        { id: "vAlign", btn: (this.valignEl = new Valign(style.valign)) },
        { id: "textWrap", btn: (this.textwrapEl = new Textwrap()) },
      ],
      [
        { id: "freeze", btn: (this.freezeEl = new Freeze()) },
        { id: "autoFilter", btn: (this.autofilterEl = new Autofilter()) },
        { id: "formula", btn: (this.formulaEl = new Formula()) },
      ],
    ];

    const { extendToolbar = {} } = data.settings;

    if (extendToolbar.left && extendToolbar.left.length > 0) {
      const btns = extendToolbar.left.map((toolbarBtn) => {
        return {
          id: toolbarBtn.id,
          btn: genBtn.call(this, toolbarBtn),
        };
      });
      this.items.unshift(btns);
    }
    if (extendToolbar.right && extendToolbar.right.length > 0) {
      const btns = extendToolbar.right.map((toolbarBtn) => {
        return {
          id: toolbarBtn.id,
          btn: genBtn.call(this, toolbarBtn),
        };
      });
      this.items.push(btns);
    }

    const cellConfigButtons = this.sheet?.options?.cellConfigButtons;
    if (cellConfigButtons?.length) {
      const configButtons = cellConfigButtons.map((config) => {
        if (config.tag) {
          return {
            id: config.id,
            btn: (this[`${config.tag}El`] = new CellConfigButtons(config)),
          };
        }
      });
      if (configButtons?.length) {
        this.items.push(configButtons);
      }
    }

    this.items.push([
      {
        id: generateUniqueId(),
        btn: (this.moreEl = new More()),
      },
    ]);

    const { disableFeatures = [] } = this.data?.settings ?? {
      disableFeatures: [],
    };

    const preparedItems = [];
    this.items.forEach((item) => {
      const filteredItems = [];
      item.forEach((subItem) => {
        if (subItem.btn) {
          if (subItem.id && disableFeatures.includes(subItem.id)) return;
          else {
            if (
              subItem.id === "import" &&
              !this.data?.settings?.allowMultipleSheets
            )
              return;
            filteredItems.push(subItem.btn);
          }
        }
      });
      if (filteredItems.length && preparedItems.length) {
        preparedItems.push(buildDivider());
      }
      preparedItems.push(filteredItems);
    });

    this.items = preparedItems;
    this.el = h("div", `${cssPrefix}-toolbar`);
    this.btns = h("div", `${cssPrefix}-toolbar-btns`);

    this.items.forEach((it) => {
      if (Array.isArray(it)) {
        it.forEach((i) => {
          this.btns.child(i.el);
          i.change = (...args) => {
            const range = this.sheet.selector.range;
            this.sheet?.trigger?.("toolbar-action", {
              action: args,
              range,
            });
            this.change(...args);
          };
        });
      } else {
        this.btns.child(it.el);
      }
    });

    this.el.child(this.btns);
    if (isHide) {
      this.el.hide();
    } else {
      this.reset();
      setTimeout(() => {
        initBtns2.call(this);
        moreResize.call(this);
      }, 0);
      bind(window, "resize", () => {
        moreResize.call(this);
      });
    }
  }

  paintformatActive() {
    return this.paintformatEl.active();
  }

  paintformatToggle() {
    this.paintformatEl.toggle();
  }

  trigger(type) {
    this[`${type}El`].click();
  }

  resetData(data) {
    this.data = data;
    this.reset();
  }

  reset() {
    if (this.isHide) return;
    
    const { data } = this;

    const { cellConfig = {} } = data;
    const { cellButtons = [] } = cellConfig;
    const style = data.getSelectedCellStyle();
    // console.log('canUndo:', data.canUndo());
    this.undoEl.setState(!data.canUndo());
    this.redoEl.setState(!data.canRedo());
    this.mergeEl.setState(data.canUnmerge(), !data.selector.multiple());
    this.autofilterEl.setState(!data.canAutofilter());
    // this.mergeEl.disabled();
    // console.log('selectedCell:', style, cell);
    const { font, format } = style;
    this.formatEl.setState(format);
    this.fontEl.setState(font.name);
    this.fontSizeEl.setState(font.size);
    this.boldEl.setState(font.bold);
    this.italicEl.setState(font.italic);
    this.underlineEl.setState(style.underline);
    this.strikeEl.setState(style.strike);
    this.textColorEl.setState(style.color);
    this.fillColorEl.setState(style.bgcolor);
    this.alignEl.setState(style.align);
    this.valignEl.setState(style.valign);
    this.textwrapEl.setState(style.textwrap);
    // console.log('freeze is Active:', data.freezeIsActive());
    this.freezeEl.setState(data.freezeIsActive());
    this.gridEl.setState(!!data.sheetConfig.gridLine);
    if (cellButtons?.length) {
      const cellMeta = data.getSelectedCellMetaData();
      cellButtons.forEach((button) => {
        const status = cellMeta[button.tag];
        this[`${button.tag}El`].setState(!!status);
      });
    }
  }
}
