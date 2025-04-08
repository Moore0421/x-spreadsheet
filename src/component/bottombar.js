import { h } from "./element";
import { bindClickoutside, unbindClickoutside } from "./event";
import { cssPrefix } from "../config";
import Icon from "./icon";
import FormInput from "./form_input";
import Dropdown from "./dropdown";
// Record: temp not used
import { xtoast } from "./message";
import { tf } from "../locale/locale";

class DropdownMore extends Dropdown {
  constructor(click) {
    const icon = new Icon("ellipsis");
    super(icon, "auto", false, "bottom-left");
    this.contentClick = click;
    
    // 修改下拉菜单的样式
    this.contentEl.css({
      "z-index": "9999",
      "position": "fixed",
      "background": "#fff",
      "box-shadow": "0 -2px 6px rgba(0,0,0,0.15)",
      "min-width": "120px",
      "max-height": "300px",
      "overflow-y": "auto",
      "border-radius": "4px"
    });

    // 添加点击事件处理
    this.headerEl.on("click", (evt) => {
      evt.stopPropagation();
      const rect = this.headerEl.el.getBoundingClientRect();
      this.contentEl.css({
        "bottom": `${window.innerHeight - rect.top + window.scrollY}px`,
        "left": `${rect.left}px`
      });
    });
  }

  reset(items) {
    if (!items || items.length === 0) {
      this.setContentChildren(
        h("div", `${cssPrefix}-item disabled`)
          .css("padding", "5px 10px")
          .css("color", "#999")
          .child("无可用选项")
      );
      return;
    }
    
    const eles = items.map((it, i) =>
      h("div", `${cssPrefix}-item`)
        .css("padding", "5px 10px")
        .css("cursor", "pointer")
        .css("width", "150px")
        .css("font-weight", "normal")
        .on("click", () => {
          this.contentClick(i);
          this.hide();
        })
        .child(it)
    );
    this.setContentChildren(...eles);
  }

  setTitle() {}
}

const menuItems = [{ key: "delete", title: tf("contextmenu.deleteSheet") }];

function buildMenuItem(item) {
  return h("div", `${cssPrefix}-item`)
    .child(item.title())
    .on("click", () => {
      this.itemClick(item.key);
      this.hide();
    });
}

function buildMenu() {
  return menuItems.map((it) => buildMenuItem.call(this, it));
}

class ContextMenu {
  constructor() {
    this.el = h("div", `${cssPrefix}-contextmenu`)
      .css("width", "160px")
      .children(...buildMenu.call(this))
      .hide();
    this.itemClick = () => {};
  }

  hide() {
    const { el } = this;
    el.hide();
    unbindClickoutside(el);
  }

  setOffset(offset) {
    const { el } = this;
    el.offset(offset);
    el.show();
    bindClickoutside(el);
  }
}

export default class Bottombar {
  constructor(
    allowMultipleSheets,
    addFunc = () => {},
    swapFunc = () => {},
    deleteFunc = () => {},
    updateFunc = () => {}
  ) {
    this.swapFunc = swapFunc;
    this.updateFunc = updateFunc;
    this.dataNames = [];
    this.activeEl = null;
    this.deleteEl = null;
    this.items = [];
    this.isInput = false;
    this.moreEl = new DropdownMore((i) => {
      console.log('DropdownMore clicked:', i);
      console.log('Current items:', this.items);
      console.log('Current dataNames:', this.dataNames);
      this.clickSwap2(this.items[i]);
    });
    this.contextMenu = new ContextMenu();
    this.contextMenu.itemClick = deleteFunc;
    const menuChildrens = allowMultipleSheets
      ? [
          new Icon("add").on("click", () => {
            addFunc();
          }),
          h("span", "").child(this.moreEl),
        ]
      : [h("span", "").child(this.moreEl)];
    
    this.arrowLeft = h("div", `${cssPrefix}-sheet-arrow left`)
      .child("◄")
      .on("click", () => this.scroll(-1));
    
    this.arrowRight = h("div", `${cssPrefix}-sheet-arrow right`)
      .child("►")
      .on("click", () => this.scroll(1));

    this.scrollEl = h("div", `${cssPrefix}-sheet-scroll`);
    
    this.el = h("div", `${cssPrefix}-bottombar`).children(
      this.contextMenu.el,
      this.arrowLeft,
      this.scrollEl.child(this.menuEl = h("ul", `${cssPrefix}-menu`).child(
        h("li", "").children(...menuChildrens)
      )),
      this.arrowRight
    );

    // 监听滚动事件以更新箭头显示状态
    this.scrollEl.on("scroll", () => this.updateArrowsVisibility());
  }

  // 添加滚动方法
  scroll(direction) {
    const { scrollEl } = this;
    const scrollLeft = scrollEl.el.scrollLeft;
    const scrollWidth = this.menuEl.el.offsetWidth;
    const clientWidth = scrollEl.el.clientWidth;
    
    // 计算每次滚动的距离(一屏的1/3)
    const scrollStep = clientWidth / 3;
    
    scrollEl.el.scrollLeft = scrollLeft + (direction * scrollStep);
  }

  // 更新箭头显示状态
  updateArrowsVisibility() {
    const { scrollEl, arrowLeft, arrowRight } = this;
    const { scrollLeft, scrollWidth, clientWidth } = scrollEl.el;
    
    // 显示/隐藏左箭头
    if(scrollLeft > 0) {
      arrowLeft.css("display", "flex");
    } else {
      arrowLeft.css("display", "none"); 
    }
    
    // 显示/隐藏右箭头
    if(scrollLeft + clientWidth < scrollWidth) {
      arrowRight.css("display", "flex");
    } else {
      arrowRight.css("display", "none");
    }
  }

  addItem(name, active, options) {
    this.dataNames.push(name);
    const item = h("li", active ? "active" : "").child(name);
    item
      .on("click", () => {
        this.clickSwap2(item);
      })
      .on("contextmenu", (evt) => {
        if (options.mode === "read") return;
        const { offsetLeft, offsetHeight } = evt.target;
        this.contextMenu.setOffset({
          left: offsetLeft,
          bottom: offsetHeight + 1,
        });
        this.deleteEl = item;
      })
      .on("dblclick", () => {
        if (options.mode === "read") return;
        if (!this.isInput) {
          this.isInput = true;
          const oldValue = item.html();
          const input = new FormInput("auto", "");
          input.val(oldValue);
          const handleInputEvent = ({ target }) => {
            const { value } = target;
            const isUnique = this.dataNames.every((name) => {
              return name.toLowerCase() !== value.toLowerCase();
            });
            if (isUnique || oldValue === value) {
              this.isInput = false;
              const nindex = this.dataNames.findIndex((it) => it === oldValue);
              this.renameItem(nindex, value);
            } else {
              xtoast(
                "Duplicate name",
                "Sheet already exists with this name",
                () => {
                  input.focus();
                }
              );
            }
          };
          input.input.on("blur", handleInputEvent);
          input.input.on("keydown", (event) => {
            const keyCode = event.keyCode || event.which;
            if (keyCode === 13) {
              input.blur();
            }
          });
          item.html("").child(input.el);
          input.focus();
        }
      });
    if (active) {
      this.clickSwap(item);
    }
    this.items.push(item);
    this.menuEl.child(item);
    this.moreEl.reset(this.dataNames);
    this.updateArrowsVisibility();
  }

  renameItem(index, value, skipSheetsTill = null) {
    this.dataNames.splice(index, 1, value);
    this.moreEl.reset(this.dataNames);
    this.items[index].html("").child(value);
    this.updateFunc(index, value, skipSheetsTill);
  }

  clear() {
    this.items.forEach((it) => {
      this.menuEl.removeChild(it.el);
    });
    this.items = [];
    this.dataNames = [];
    this.moreEl.reset(this.dataNames);
    this.updateArrowsVisibility(); 
  }

  deleteItem() {
    return [-1];
  }

  clickSwap2(item) {
    const index = this.items.findIndex((it) => it === item);
    this.clickSwap(item);
    this.activeEl.toggle();
    this.swapFunc(index);
    
    // 重置滚动位置
    const container = document.querySelector(`.x-spreadsheet-sheet`);
    if (container) {
      container.scrollLeft = 0;
      container.scrollTop = 0;
    }
  }

  clickSwap(item) {
    if (this.activeEl !== null) {
      this.activeEl.toggle();
    }
    this.activeEl = item;
  }
}
