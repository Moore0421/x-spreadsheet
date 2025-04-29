import { h } from "./element";
import { cssPrefix } from "../config";
import { bindClickoutside, unbindClickoutside } from "./event";

export default class Datepicker {
  constructor(options = {}) {
    this.options = options;
    this.width = options.width || 280;
    this.height = options.height || 350;

    // 当前编辑的单元格信息
    this.currentCell = null;
    this.currentRi = null;
    this.currentCi = null;

    // 当前选择的日期
    this.value = null;

    // 创建主容器
    this.el = h("div", `${cssPrefix}-datepicker`)
      .css("width", `${this.width}px`)
      .css("height", `${this.height}px`)
      .css("position", "fixed")
      .css("background", "#fff")
      .css("box-shadow", "0 2px 12px rgba(0,0,0,0.15)")
      .css("z-index", "10000")
      .css("border-radius", "4px")
      .hide();

    // 创建标题栏
    this.header = h("div", `${cssPrefix}-datepicker-header`)
      .css("padding", "10px")
      .css("border-bottom", "1px solid #e0e0e0")
      .css("display", "flex")
      .css("justify-content", "space-between")
      .css('cursor', 'move')
      .css("align-items", "center");

    this.title = h("div", `${cssPrefix}-datepicker-title`)
      .css("font-weight", "bold")
      .html("选择日期");

    this.closeBtn = h("div", `${cssPrefix}-datepicker-close`)
      .css("cursor", "pointer")
      .css("font-size", "16px")
      .html("×")
      .on("click", () => this.hide());

    this.header.children(this.title, this.closeBtn);

    // 创建日历内容区域
    this.content = h("div", `${cssPrefix}-datepicker-content`)
      .css("padding", "10px")
      .css("height", "calc(100% - 110px)")
      .css("overflow", "auto");

    // 创建日历组件
    this.calendar = this.createCalendar();
    this.content.child(this.calendar);

    // 创建底部按钮区域
    this.footer = h("div", `${cssPrefix}-datepicker-footer`)
      .css("padding", "10px")
      .css("border-top", "1px solid #e0e0e0")
      .css("display", "flex")
      .css("justify-content", "flex-end");

    this.cancelBtn = h(
      "button",
      `${cssPrefix}-datepicker-button ${cssPrefix}-datepicker-cancel`
    )
      .css("padding", "5px 15px")
      .css("margin-right", "10px")
      .css("border", "1px solid #ddd")
      .css("background", "#f5f5f5")
      .css("border-radius", "4px")
      .css("cursor", "pointer")
      .html("取消")
      .on("click", () => this.hide());

    this.confirmBtn = h(
      "button",
      `${cssPrefix}-datepicker-button ${cssPrefix}-datepicker-confirm`
    )
      .css("padding", "5px 15px")
      .css("border", "none")
      .css("background", "#4b89ff")
      .css("color", "#fff")
      .css("border-radius", "4px")
      .css("cursor", "pointer")
      .html("确认")
      .on("click", () => {
        if (this._changeCallback && this.value) {
          // 格式化为yyyy-mm-dd
          const y = this.value.getFullYear();
          const m = String(this.value.getMonth() + 1).padStart(2, "0");
          const d = String(this.value.getDate()).padStart(2, "0");
          this._changeCallback(`${y}-${m}-${d}`);
        }
        this.hide();
      });

    this.footer.children(this.cancelBtn, this.confirmBtn);

    // 组装所有部分
    this.el.children(this.header, this.content, this.footer);

    // 回调函数 - 重要：这里使用属性而不是方法
    this._changeCallback = () => {};
    // 添加拖拽功能
    this.initDrag();
  }

  // 初始化拖拽功能
  initDrag() {
    let isDragging = false;
    let startX, startY, initialX, initialY;

    this.header.on("mousedown", (e) => {
      if (e.target === this.closeBtn.el) return;
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

  // 创建简单的日历组件
  createCalendar() {
    const now = new Date();
    this.currentYear = now.getFullYear();
    this.currentMonth = now.getMonth();
    this.value = now; // 默认选中今天

    const calendarEl = h("div", `${cssPrefix}-datepicker-calendar`)
      .css("width", "100%")
      .css("user-select", "none");

    // 创建年月选择器
    const yearMonthSelector = h("div", `${cssPrefix}-datepicker-year-month`)
      .css("display", "flex")
      .css("justify-content", "space-between")
      .css("margin-bottom", "10px")
      .css("align-items", "center");

    // 上一月按钮
    const prevMonthBtn = h("div", `${cssPrefix}-datepicker-prev`)
      .css("cursor", "pointer")
      .css("padding", "0 10px")
      .html("&lt;")
      .on("click", () => {
        this.currentMonth -= 1;
        if (this.currentMonth < 0) {
          this.currentMonth = 11;
          this.currentYear -= 1;
        }
        this.updateCalendar();
      });

    // 年月显示
    this.yearMonthText = h(
      "div",
      `${cssPrefix}-datepicker-year-month-text`
    ).css("font-weight", "bold");

    // 下一月按钮
    const nextMonthBtn = h("div", `${cssPrefix}-datepicker-next`)
      .css("cursor", "pointer")
      .css("padding", "0 10px")
      .html("&gt;")
      .on("click", () => {
        this.currentMonth += 1;
        if (this.currentMonth > 11) {
          this.currentMonth = 0;
          this.currentYear += 1;
        }
        this.updateCalendar();
      });

    yearMonthSelector.children(prevMonthBtn, this.yearMonthText, nextMonthBtn);

    // 创建星期头部
    const weekHeader = h("div", `${cssPrefix}-datepicker-week-header`)
      .css("display", "grid")
      .css("grid-template-columns", "repeat(7, 1fr)")
      .css("text-align", "center")
      .css("font-weight", "bold")
      .css("margin-bottom", "5px");

    const weekdays = ["日", "一", "二", "三", "四", "五", "六"];
    weekdays.forEach((day) => {
      weekHeader.child(h("div").html(day));
    });

    // 创建日期网格
    this.daysGrid = h("div", `${cssPrefix}-datepicker-days`)
      .css("display", "grid")
      .css("grid-template-columns", "repeat(7, 1fr)")
      .css("grid-gap", "5px")
      .css("text-align", "center");

    calendarEl.children(yearMonthSelector, weekHeader, this.daysGrid);

    // 初始化日历
    this.updateCalendar();

    return calendarEl;
  }

  // 更新日历显示
  updateCalendar() {
    // 更新年月显示
    const monthNames = [
      "一月",
      "二月",
      "三月",
      "四月",
      "五月",
      "六月",
      "七月",
      "八月",
      "九月",
      "十月",
      "十一月",
      "十二月",
    ];
    this.yearMonthText.html(
      `${this.currentYear}年 ${monthNames[this.currentMonth]}`
    );

    // 清空日期网格
    this.daysGrid.html("");

    // 获取当月第一天是星期几
    const firstDay = new Date(this.currentYear, this.currentMonth, 1).getDay();

    // 获取当月天数
    const daysInMonth = new Date(
      this.currentYear,
      this.currentMonth + 1,
      0
    ).getDate();

    // 添加空白格子
    for (let i = 0; i < firstDay; i++) {
      this.daysGrid.child(h("div"));
    }

    // 添加日期格子
    for (let day = 1; day <= daysInMonth; day++) {
      const dayEl = h("div", `${cssPrefix}-datepicker-day`)
        .css("cursor", "pointer")
        .css("padding", "5px")
        .css("border-radius", "4px")
        .html(day);

      // 检查是否是当前选中的日期
      if (
        this.value &&
        this.value.getFullYear() === this.currentYear &&
        this.value.getMonth() === this.currentMonth &&
        this.value.getDate() === day
      ) {
        dayEl.css("background", "#4b89ff").css("color", "#fff");
      }

      // 点击日期
      dayEl.on("click", () => {
        this.value = new Date(this.currentYear, this.currentMonth, day);
        this.updateCalendar();
      });

      this.daysGrid.child(dayEl);
    }
  }

  // 设置日期
  setValue(date) {
    if (date instanceof Date) {
      this.value = date;
    } else if (typeof date === "string") {
      // 尝试解析日期字符串
      try {
        this.value = new Date(date);
        this.currentYear = this.value.getFullYear();
        this.currentMonth = this.value.getMonth();
      } catch (e) {
        console.error("无法解析日期:", e);
        this.value = new Date();
      }
    } else {
      this.value = new Date();
    }

    this.updateCalendar();
    return this;
  }

  // 设置当前编辑的单元格信息
  setCell(ri, ci, cell) {
    this.currentRi = ri;
    this.currentCi = ci;
    this.currentCell = cell;

    // 如果单元格有日期值，设置为当前值
    if (cell && cell.text) {
      this.setValue(cell.text);
    } else {
      this.setValue(new Date());
    }

    return this;
  }

  // 设置回调函数 - 重要：这是一个方法，用于设置回调属性
  setChange(callback) {
    if (typeof callback === "function") {
      this._changeCallback = callback;
    }
    return this;
  }

  // 设置位置
  setOffset(offset) {
    if (offset) {
      // 确保弹窗显示在视口内
      const maxX = window.innerWidth - this.el.el.offsetWidth;
      const maxY = window.innerHeight - this.el.el.offsetHeight;

      const x = Math.max(0, Math.min(offset.left, maxX));
      const y = Math.max(0, Math.min(offset.top, maxY));

      this.el
        .css("transform", "none")
        .css("left", `${x}px`)
        .css("top", `${y}px`);
    } else {
      // 默认居中显示
      this.el
        .css("transform", "translate(-50%, -50%)")
        .css("left", "50%")
        .css("top", "50%");
    }
    return this;
  }

  // 显示日期选择器
  show() {
    this.setOffset(); // 默认居中显示
    this.el.show();

    // 绑定点击外部事件，但排除日期选择器内部元素
    bindClickoutside(this.el, () => {
      if (this.el.visible()) {
        this.hide();
      }
    });

    return this;
  }

  // 隐藏日期选择器
  hide() {
    this.el.hide();
    unbindClickoutside(this.el);
    return this;
  }

  change(callback) {
    this._changeCallback = callback;
    return this;
  }
}
