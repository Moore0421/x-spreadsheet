import { h } from "./element";
import { cssPrefix } from "../config";

export default class PopupDialog {
  constructor(options = {}) {
    this.options = options;
    this.title = null;
    this.width = null;
    this.height = null;

    // 创建弹窗元素
    this.el = h("div", `${cssPrefix}-popup-dialog`)
      .css("width", `${this.width}px`)
      .css("height", `${this.height}px`)
      .css("position", "fixed")
      .css("background", "#fff")
      .css("box-shadow", "0 2px 12px rgba(0,0,0,0.15)")
      .css("z-index", "10000")
      .css("border-radius", "4px")
      .css("transform", "translate(-50%, -50%)")
      .hide();

    // 创建弹窗头部
    this.header = h("div", `${cssPrefix}-popup-header`)
      .css("padding", "10px")
      .css("border-bottom", "1px solid #e0e0e0")
      .css("display", "flex")
      .css("justify-content", "space-between")
      .css("align-items", "center")
      .css('cursor', 'move')
      .child(h("div", `${cssPrefix}-popup-title`).html(this.title))
      .child(
        h("div", `${cssPrefix}-popup-close`)
          .html("×")
          .css("cursor", "pointer")
          .css("font-size", "16px")
          .on("click", () => this.hide())
      );

    // 创建弹窗内容区域
    this.content = h("div", `${cssPrefix}-popup-content`)
      .css("padding", "10px")
      .css("height", "calc(100% - 110px)")
      .css("overflow", "auto");

    // 创建弹窗底部
    this.footer = h("div", `${cssPrefix}-popup-footer`)
      .css("padding", "10px")
      .css("border-top", "1px solid #e0e0e0")
      .css("display", "flex")
      .css("justify-content", "flex-end")
      .child(
        h("button", `${cssPrefix}-popup-button ${cssPrefix}-popup-cancel`)
          .css("padding", "5px 15px")
          .css("margin-right", "10px")
          .css("border", "1px solid #ddd")
          .css("background", "#f5f5f5")
          .html("取消")
          .on("click", () => this.hide())
      )
      .child(
        h("button", `${cssPrefix}-popup-button ${cssPrefix}-popup-confirm`)
          .css("padding", "5px 15px")
          .css("border", "1px solid #ddd")
          .html("确定")
          .on("click", () => {
            if (this.onConfirm) this.onConfirm();
            this.hide();
          })
      );

    // 组装弹窗
    this.el.children(this.header, this.content, this.footer);

    // 添加到body
    document.body.appendChild(this.el.el);

    // 添加拖拽功能
    this.initDrag();
  }

  // 初始化拖拽功能
  initDrag() {
    let isDragging = false;
    let startX, startY, initialX, initialY;

    this.header.on("mousedown", (e) => {
      isDragging = true;
      startX = e.clientX;
      startY = e.clientY;
      const rect = this.el.el.getBoundingClientRect();
      initialX = rect.left;
      initialY = rect.top;

      const mousemove = (e) => {
        if (!isDragging) return;
        e.preventDefault();

        const deltaX = e.clientX - startX;
        const deltaY = e.clientY - startY;

        let newX = initialX + deltaX;
        let newY = initialY + deltaY;

        // 确保弹窗不会超出视口
        const maxX = window.innerWidth - this.el.el.offsetWidth;
        const maxY = window.innerHeight - this.el.el.offsetHeight;

        newX = Math.max(0, Math.min(newX, maxX));
        newY = Math.max(0, Math.min(newY, maxY));

        this.el
          .css("transform", "none")
          .css("left", `${newX}px`)
          .css("top", `${newY}px`);
      };

      const mouseup = () => {
        isDragging = false;
        document.removeEventListener("mousemove", mousemove);
        document.removeEventListener("mouseup", mouseup);
      };

      document.addEventListener("mousemove", mousemove);
      document.addEventListener("mouseup", mouseup);
    });
  }

  // 设置内容
  setContent(content) {
    if (typeof content === "string") {
      this.content.html(content);
    } else {
      this.content.html("");
      this.content.child(content);
    }
    return this;
  }

  // 设置确认回调
  onConfirm(callback) {
    this.onConfirm = callback;
    return this;
  }

  // 显示弹窗
  show() {
    this.el.show();
    // 添加滚动条
    this.content.css("overflow-y", "auto");
    return this;
  }

  // 隐藏弹窗
  hide() {
    this.el.hide();
    return this;
  }

  // 销毁弹窗
  destroy() {
    document.body.removeChild(this.el.el);
  }

  // 添加 setTitle 方法，确保标题正确应用
  setTitle(title) {
    this.title = title;
    this.header.el.querySelector(`.${cssPrefix}-popup-title`).innerHTML = title;
    return this;
  }

  // 添加 setSize 方法，确保宽度和高度正确应用
  setSize(width, height) {
    this.width = width;
    this.height = height;
    this.el.css("width", `${this.width}px`).css("height", `${this.height}px`);
    return this;
  }
}
