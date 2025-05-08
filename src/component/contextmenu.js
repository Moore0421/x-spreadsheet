import { h } from "./element";
import { bindClickoutside, unbindClickoutside } from "./event";
import { cssPrefix } from "../config";
import { tf } from "../locale/locale";
import { CELL_REF_REGEX, SHEET_TO_CELL_REF_REGEX } from "../constants";

const menuItems = [
  {
    key: "navigate",
    title: tf("contextmenu.navigate"),
    label: "Ctrl+[",
  },
  { key: "copy", title: tf("contextmenu.copy"), label: "Ctrl+C" },
  { key: "cut", title: tf("contextmenu.cut"), label: "Ctrl+X" },
  { key: "paste", title: tf("contextmenu.paste"), label: "Ctrl+V" },
  {
    key: "paste-value",
    title: tf("contextmenu.pasteValue"),
    label: "Ctrl+Shift+V",
  },
  {
    key: "paste-format",
    title: tf("contextmenu.pasteFormat"),
    label: "Ctrl+Alt+V",
  },
  { key: "divider" },
  { key: "insert-row", title: tf("contextmenu.insertRow") },
  { key: "insert-column", title: tf("contextmenu.insertColumn") },
  { key: "divider" },
  { key: "delete-row", title: tf("contextmenu.deleteRow") },
  { key: "delete-column", title: tf("contextmenu.deleteColumn") },
  { key: "delete-cell-text", title: tf("contextmenu.deleteCellText") },
  { key: "hide", title: tf("contextmenu.hide") },
  { key: "divider" },
  { key: "validation", title: tf("contextmenu.validation") },
  { key: "divider" },
  { key: "cell-type", title: tf("contextmenu.cellType"), subMenus: [
    { key: "cell-type-text", title: tf("contextmenu.cellTypeText") },
    { key: "cell-type-date", title: tf("contextmenu.cellTypeDate") },
    { key: "cell-type-tree", title: tf("contextmenu.cellTypeTree") },
    { key: "cell-type-popup", title: tf("contextmenu.cellTypePopup") },
  ]},
  { key: "divider" },
  { key: "cell-editable", title: tf("contextmenu.celleditable") },
  { key: "cell-non-editable", title: tf("contextmenu.cellnoneditable") },
  { key: "divider" },
  { key: "set-row-id", title: tf("contextmenu.setRowId") },
  { key: "set-data-cell", title: tf("contextmenu.setDataCell") },
  { key: "cancel-data-cell", title: tf("contextmenu.cancelDataCell") },
  { key: "divider" },
  { key: "set-data-list", title: tf("contextmenu.setDataList") },
  { key: "cancel-data-list", title: tf("contextmenu.cancelDataList") }
];

function buildSubMenuItems(items) {
  const subMenus = items.map((it) => buildMenuItem.call(this, it));
  return subMenus;
}

function buildMenuItem(item) {
  if (item.key === "divider") {
    return h("div", `${cssPrefix}-item divider`);
  }
  const ele = h("div", `${cssPrefix}-item`)
    .on("click", () => {
      const range = this.sheet.selector.range;
      if (item.key === "set-data-list") {
        this.sheet.data.rows.setDataList(range);
      } else if (item.key === "cancel-data-list") {
        this.sheet.data.rows.cancelDataList();
      } else if (item.key === "set-row-id") {
        this.sheet.data.rows.handleSetRowId();
      } else {
        this.sheet?.trigger?.("context-menu-action", {
          action: [item.key],
          range,
        });
      }
      this.itemClick(item.key);
      this.hide();
    })
    .children(
      typeof item.title === "function" ? item.title() : (item.title ?? ""),
      h("div", "label").child(item.label || "")
    )
    .attr("data-key", item.key);

  if (item.subMenus) {
    const arrowIcon = h("div", `${cssPrefix}-icon label `).child(
      h("div", `${cssPrefix}-icon-img arrow-right`)
    );
    ele.child(arrowIcon);
    const subMenus = buildSubMenuItems.call(this, item.subMenus);
    const subMenuEl = h("div", `${cssPrefix}-context-sub-menu`)
      .children(...subMenus)
      .hide();
    ele
      .child(subMenuEl)
      .on("mouseover", () => {
        const rect = ele.box();
        const subMenuElBox = subMenuEl.box();
        const view = this.viewFn();
        const totalWidth = subMenuElBox.width + rect.x + rect.width;
        let left = rect.left;
        let top = rect.top;
        if (totalWidth + rect.width > view.width) {
          left = left - 200;
        } else {
          left = left + 200;
        }
        const subMenuHeight = item.subMenus.length * 50;
        const totalHeight = subMenuHeight + rect.y + rect.height;
        if (totalHeight + subMenuHeight > view.height) {
          top = rect.top - subMenuHeight / 2;
        }
        subMenuEl.css("left", `${left}px`);
        subMenuEl.css("top", `${top}px`);
        subMenuEl.show();
      })
      .on("mouseout", () => {
        subMenuEl.hide();
      });
  }

  return ele;
}

function checkCommentButtonStatus(element, key) {
  const { sheet = {} } = this;
  const { data = {} } = sheet;
  const cell = data.getSelectedCell();
  const { c = [] } = cell ?? {};
  if (c?.length) {
    if (key === "add-comment") {
      element.hide();
    } else if (key === "show-comment") {
      element.show();
    }
  } else {
    if (key === "add-comment") {
      element.show();
    } else if (key === "show-comment") {
      element.hide();
    }
  }
}

function setNavigateVisibility(element) {
  const { sheet = {} } = this;
  const { data = {} } = sheet;
  const cell = data.getSelectedCell();
  const { f } = cell ?? {};
  if (f?.match(SHEET_TO_CELL_REF_REGEX) || f?.match(CELL_REF_REGEX)) {
    element.show();
  } else {
    element.hide();
  }
}

function handleDynamicMenu() {
  const { menuItems, extendedContextMenu, sheet } = this;
  menuItems.forEach((element) => {
    const key = element.attr("data-key");
    if (key === "add-comment" || key === "show-comment") {
      checkCommentButtonStatus.call(this, element, key);
    } else if (key === "navigate") {
      setNavigateVisibility.call(this, element);
    } else if (extendedContextMenu?.length) {
      const match = extendedContextMenu?.find((menu) => menu.key === key);
      if (match?.visibility) {
        const status = match?.visibility?.(sheet, key);
        if (status) {
          element.show();
        } else {
          element.hide();
        }
      }
    }
  });
  const setDataListEl = this.menuItems.find(ele => ele.attr("data-key") === "set-data-list");
  const cancelDataListEl = this.menuItems.find(ele => ele.attr("data-key") === "cancel-data-list");
  if (this.sheet.data.settings.dataListRange) {
    setDataListEl && setDataListEl.hide();
    cancelDataListEl && cancelDataListEl.show();
  } else {
    setDataListEl && setDataListEl.show();
    cancelDataListEl && cancelDataListEl.hide();
  }
}

function buildMenu() {
  const { sheet, extendedContextMenu } = this;
  let menu = [];
  if (typeof sheet?.options?.comment === "object") {
    menu = [
      {
        key: "add-comment",
        title: tf("contextmenu.addComment"),
      },
      {
        key: "show-comment",
        title: tf("contextmenu.showComment"),
      },
      { key: "divider" },
      ...menuItems,
    ];
  } else {
    menu = menuItems;
  }
  const buildInMenus = menu.map((it) => buildMenuItem.call(this, it));
  let additionalMenus = [];
  if (extendedContextMenu?.length) {
    const extMenu = Array.from(extendedContextMenu);
    extMenu.push({ key: "divider" });
    additionalMenus = extMenu.map((it) => buildMenuItem.call(this, it));
  }
  return [...(additionalMenus ?? []), ...buildInMenus];
}

function addCustomScrollbar(container) {
  const scrollbarThumb = h("div", `${cssPrefix}-custom-scrollbar-thumb`);
  const scrollbarTrack = h("div", `${cssPrefix}-custom-scrollbar-track`).child(
    scrollbarThumb
  );
  container.appendChild(scrollbarTrack.el);

  const updateScrollbar = () => {
    const viewableRatio = container.clientHeight / container.scrollHeight;
    const thumbHeight = Math.max(viewableRatio * container.clientHeight, 30);
    scrollbarThumb.css("height", `${thumbHeight}px`)
    scrollbarTrack.css(
      "display",
      container.scrollHeight <= container.clientHeight ? "none" : "block"
    );
  };

  let isDragging = false;
  let startY, startScrollTop;

  const syncThumbPosition = () => {
    const maxScroll = container.scrollHeight - container.clientHeight;
    const maxThumbTravel = container.clientHeight - scrollbarThumb.el.offsetHeight;
    const scrollRatio = container.scrollTop / maxScroll;
    const thumbOffset = scrollRatio * maxThumbTravel;
    scrollbarThumb.css('transform', `translateY(${thumbOffset}px)`)
    scrollbarTrack.css('transform', `translateY(${container.scrollTop}px)`)
  };

  const handleMouseMove = (e) => {
    if (!isDragging) return;
    const deltaY = e.clientY - startY;
    const maxScroll = container.scrollHeight - container.clientHeight;
    const maxThumbTravel = container.clientHeight - scrollbarThumb.el.offsetHeight;
    const newThumbPos = Math.max(
      0,
      Math.min(startScrollTop + deltaY, maxThumbTravel)
    );
    const scrollRatio = newThumbPos / maxThumbTravel;
    container.scrollTop = scrollRatio * maxScroll;
  };

  const handleMouseUp = () => {
    isDragging = false;
    document.removeEventListener("mousemove", handleMouseMove);
    document.removeEventListener("mouseup", handleMouseUp);
    setTimeout(() => {
      window.isCustomThumbDrag = null;
    });
  };

  scrollbarThumb.on("mousedown", (e) => {
    isDragging = true;
    window.isCustomThumbDrag = true;
    startY = e.clientY;
    startScrollTop =
      parseFloat(
        scrollbarThumb.el.style.transform
          .replace("translateY(", "")
          .replace("px)", "")
      ) || 0;

    document.addEventListener("mousemove", handleMouseMove);
    document.addEventListener("mouseup", handleMouseUp);
    e.preventDefault();
  });

  container.addEventListener(
    "wheel",
    (e) => {
      e.preventDefault();
      container.scrollTop += e.deltaY * 0.5;
    },
    { passive: false }
  );

  container.addEventListener("scroll", syncThumbPosition);
  const resizeObserver = new ResizeObserver(() => {
    updateScrollbar();
    syncThumbPosition();
  });
  resizeObserver.observe(container);
  updateScrollbar();
}

export default class ContextMenu {
  constructor(sheetContext, viewFn, isHide = false, extendedContextMenu = []) {
    this.extendedContextMenu = extendedContextMenu;
    this.sheet = sheetContext;
    this.menuItems = buildMenu.call(this);
    this.el = h("div", `${cssPrefix}-contextmenu`)
      .children(...this.menuItems)
      .hide();
    this.viewFn = viewFn;
    this.itemClick = () => {};
    this.isHide = isHide;
    this.setMode("range");
    this.lastCoordinate = { x: 0, y: 0 };
    addCustomScrollbar(this.el.el);
  }

  // row-col: the whole rows or the whole cols
  // range: select range
  //This may cause any issue in future
  setMode(mode) {
    const hideEl = this.menuItems.find((ele) => {
      return ele.attr("data-key") === "hide";
    });
    if (hideEl) {
      if (mode === "row-col") {
        hideEl.show();
      } else {
        hideEl.hide();
      }
    }
  }

  hide() {
    const { el } = this;
    el.hide();
    unbindClickoutside(el);
  }

  show(evt) {
    // 在预览模式下禁用右键菜单
    if (this.sheet?.data?.settings?.mode === 'preview' || this.sheet?.data?.settings?.mode === 'enabled') {
      return;
    }
    
    const { sheet } = this;
    const { selector } = sheet.data;
    const cell = sheet.data.getCell(sheet.data.selector.ri, sheet.data.selector.ci);
    
    // 如果单元格不可编辑且不是设计模式，隐藏右键菜单
    if (cell && cell.editable === false && sheet.data.settings.mode !== 'design') {
      this.hide();
      return;
    }

    this.lastCoordinate = { x: evt.offsetX, y: evt.offsetY };
    if (this.isHide) return;
    const { el } = this;
    handleDynamicMenu.call(this);
    const { width } = el.show().offset();
    const view = this.viewFn();
    const vhf = view.height / 2;
    let left = evt.offsetX;
    if (view.width - evt.offsetX <= width) {
      left -= width;
    }
    el.css("left", `${left}px`);
    if (evt.offsetY > vhf) {
      el.css("bottom", `${view.height - evt.offsetY}px`)
        .css("max-height", `${evt.offsetY}px`)
        .css("top", "auto");
    } else {
      el.css("top", `${evt.offsetY}px`)
        .css("max-height", `${view.height - evt.offsetY}px`)
        .css("bottom", "auto");
    }
    bindClickoutside(el);
    
    // 只在选中整行时显示设置行ID选项
    const setRowIdEl = this.menuItems.find(it => it.attr("data-key") === "set-row-id");
    if (setRowIdEl) {
      const { sri, eri, sci, eci } = selector.range;
      // 检查是否只选中了一行
      const isWholeRow = sci === 0 && eci === sheet.data.cols.len - 1 && sri === eri;
      if (isWholeRow) {
        setRowIdEl.show();
      } else {
        setRowIdEl.hide();
      }
    }
  }
}
