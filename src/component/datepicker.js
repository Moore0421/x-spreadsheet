import Calendar from "./calendar";
import { h } from "./element";
import { cssPrefix } from "../config";
import { bindClickoutside, unbindClickoutside } from './event';

export default class Datepicker {
  constructor() {
    // 创建主容器
    this.el = h("div", `${cssPrefix}-datepicker`)
      .css('position', 'fixed')
      .css('background', '#fff')
      .css('box-shadow', '0 2px 12px rgba(0,0,0,0.15)')
      .css('border-radius', '4px')
      .css('z-index', '10000')
      .hide();

    // 创建日历组件
    this.calendar = new Calendar(new Date());

    // 创建内容区域
    this.content = h('div', `${cssPrefix}-datepicker-content`)
      .css('padding', '10px')
      .child(this.calendar.el);

    // 创建底部按钮区域
    this.footer = h('div', `${cssPrefix}-datepicker-footer`)
      .css('padding', '8px')
      .css('text-align', 'right')
      .css('border-top', '1px solid #e0e0e0');

    // 创建取消按钮
    this.cancelBtn = h('button', `${cssPrefix}-datepicker-btn ${cssPrefix}-datepicker-cancel`)
      .css('padding', '5px 15px')
      .css('margin-right', '8px')
      .css('border', '1px solid #ddd')
      .css('background', '#f5f5f5')
      .css('border-radius', '4px')
      .css('cursor', 'pointer')
      .html('取消')
      .on('click', () => this.hide());

    // 创建确认按钮
    this.confirmBtn = h('button', `${cssPrefix}-datepicker-btn ${cssPrefix}-datepicker-confirm`)
      .css('padding', '5px 15px')
      .css('border', 'none')
      .css('background', '#4b89ff')
      .css('color', '#fff')
      .css('border-radius', '4px')
      .css('cursor', 'pointer')
      .html('确认')
      .on('click', () => {
        if (this.changeCallback) {
          this.changeCallback(this.value);
        }
        this.hide();
      });

    this.footer.children(this.cancelBtn, this.confirmBtn);

    // 组装所有部分
    this.el.children(this.content, this.footer);

    // 监听日历选择变化
    this.calendar.selectChange = (date) => {
      this.value = date;
    };
  }

  setValue(date) {
    if (typeof date === "string") {
      if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(date)) {
        const dateObj = new Date(date.replace(new RegExp("-", "g"), "/"));
        if (!isNaN(dateObj.getTime())) {
          this.value = dateObj;
          this.calendar.setValue(dateObj);
        }
      }
    } else if (date instanceof Date) {
      this.value = date;
      this.calendar.setValue(date);
    }
    return this;
  }

  setChange(cb) {
    // 存储回调函数而不是立即执行
    this.changeCallback = cb;
    return this;
  }

  show() {
    this.el.show();
    bindClickoutside(this.el);
    return this;
  }

  hide() {
    this.el.hide();
    unbindClickoutside(this.el);
    return this;
  }

  change(date) {
    if (this.changeCallback) {
      this.changeCallback(date);
    }
  }
  
  setOffset(offset) {
    if (offset) {
      // 确保弹窗显示在视口内
      const maxX = window.innerWidth - this.el.el.offsetWidth;
      const maxY = window.innerHeight - this.el.el.offsetHeight;
      
      const x = Math.max(0, Math.min(offset.left, maxX));
      const y = Math.max(0, Math.min(offset.top, maxY));
      
      this.el
        .css('left', `${x}px`)
        .css('top', `${y}px`);
    }
    return this;
  }
}
