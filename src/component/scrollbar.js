import { h } from "./element";
import { cssPrefix } from "../config";

export default class Scrollbar {
  constructor(vertical) {
    this.vertical = vertical;
    this.moveFn = null;
    this.el = h(
      "div",
      `${cssPrefix}-scrollbar ${vertical ? "vertical" : "horizontal"}`,
    )
      .child((this.contentEl = h("div", "")))
      .on("mousemove.stop", () => {})
      .on("scroll.stop", (evt) => {
        const { scrollTop, scrollLeft } = evt.target;
        if (this.moveFn) {
          this.moveFn(this.vertical ? scrollTop : scrollLeft, evt);
        }
      });
  }

  move(v) {
    this.el.scroll(v);
    return this;
  }

  scroll() {
    return this.el.scroll();
  }

  set(distance, contentDistance) {
    const d = distance - 1;
    if (contentDistance > d) {
      const cssKey = this.vertical ? "height" : "width";
      this.el.css(cssKey, `${d - 15}px`).show();
      this.contentEl
        .css(this.vertical ? "width" : "height", "1px")
        .css(cssKey, `${contentDistance}px`);
    } else {
      this.el.hide();
    }
    return this;
  }
}
