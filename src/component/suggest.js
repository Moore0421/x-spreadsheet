import { h } from "./element";
import { bindClickoutside, unbindClickoutside } from "./event";
import { cssPrefix } from "../config";
import { isElementInView } from "../utils";

function inputMovePrev(evt) {
  evt.preventDefault();
  evt.stopPropagation();
  const { filterItems, el } = this;
  if (filterItems.length <= 0) return;
  if (this.itemIndex >= 0) filterItems[this.itemIndex].toggle();
  this.itemIndex -= 1;
  if (this.itemIndex < 0) {
    this.itemIndex = filterItems.length - 1;
  }
  filterItems[this.itemIndex].toggle();
  if (!isElementInView(el?.el, filterItems[this.itemIndex].el))
    el?.el?.scrollBy?.({
      top: -120,
      behavior: "smooth",
    });
}

function inputMoveNext(evt) {
  evt.stopPropagation();
  const { filterItems, el } = this;
  if (filterItems.length <= 0) return;
  if (this.itemIndex >= 0) filterItems[this.itemIndex].toggle();
  this.itemIndex += 1;
  if (this.itemIndex > filterItems.length - 1) {
    this.itemIndex = 0;
  }
  filterItems[this.itemIndex].toggle();
  if (!isElementInView(el?.el, filterItems[this.itemIndex].el))
    el?.el?.scrollBy?.({
      top: 120,
      behavior: "smooth",
    });
}

function inputEnter(evt) {
  evt.preventDefault();
  const { filterItems } = this;
  if (filterItems.length <= 0) return;
  evt.stopPropagation();
  if (this.itemIndex < 0) this.itemIndex = 0;
  filterItems[this.itemIndex].el.click();
  this.hide();
}

function inputKeydownHandler(evt) {
  const { keyCode } = evt;
  if (evt.ctrlKey) {
    evt.stopPropagation();
  }
  switch (keyCode) {
    case 37: // left
      evt.stopPropagation();
      break;
    case 38: // up
      inputMovePrev.call(this, evt);
      break;
    case 39: // right
      evt.stopPropagation();
      break;
    case 40: // down
      inputMoveNext.call(this, evt);
      break;
    case 13: // enter
      inputEnter.call(this, evt);
      break;
    case 9:
      inputEnter.call(this, evt);
      break;
    default:
      evt.stopPropagation();
      break;
  }
}

export default class Suggest {
  constructor(
    items,
    itemClick,
    isMultiSelect = false,
    selectedItems,
    width = "200px"
  ) {
    this.filterItems = [];
    this.items = items;
    this.el = h("div", `${cssPrefix}-suggest`)
      .attr("tabindex", "0")
      .css("width", width)
      .css("max-height", "400px")
      .css("overflow", "auto")
      .hide();
    this.itemClick = itemClick;
    this.itemIndex = -1;
    this.isMultiSelect = isMultiSelect;
    this.selectedItems = selectedItems;
  }

  setOffset(v) {
    this.el.cssRemoveKeys("top", "bottom").offset(v);
  }

  hide() {
    const { el } = this;
    this.filterItems = [];
    this.itemIndex = -1;
    el.hide();
    unbindClickoutside(this.el.parent());
  }

  setItems(items) {
    this.items = items;
    // this.search('');
  }

  setSelectedItems(items) {
    this.selectedItems = items;
  }

  search(word) {
    let { items } = this;
    if (!/^\s*$/.test(word)) {
      items = items.filter((it) =>
        (it.key || it).startsWith(word.toUpperCase())
      );
    }
    items = items.map((it) => {
      let { title } = it;
      if (title) {
        if (typeof title === "function") {
          title = title();
        }
      } else {
        title = it;
      }
      if (this.isMultiSelect) {
        const checkBoxInput = h("input", `${cssPrefix}-item-checkbox`).css(
          "margin-right",
          "8px"
        );
        checkBoxInput
          .attr("type", "checkbox")
          .attr("id", it.key)
          .attr("name", it.key)
          .attr("value", it.key);

        if (this.selectedItems?.includes?.(it.key))
          checkBoxInput.attr("checked", true);

        const inputLabel = h("label", `${cssPrefix}-item-label`);
        inputLabel.attr("for", it.key);
        inputLabel.child(title);

        checkBoxInput.on("click", (event) => {
          event.stopPropagation();
          this.itemClick(it);
        });

        inputLabel.on("click", (event) => {
          event.stopPropagation();
        });

        const item = h("div", `${cssPrefix}-item`)
          .css("display", "flex")
          .children(checkBoxInput, inputLabel)
          .on("click.stop", () => {
            inputLabel.el.click();
          });
        return item;
      } else {
        const item = h("div", `${cssPrefix}-item`)
          .css("height", "unset")
          .css("overflow-wrap", "break-word")
          .css("border-bottom", "0.1px solid #EBEBEB")
          .child(title)
          .on("click.stop", () => {
            this.itemClick(it);
            this.hide();
          });
        if (it.label) {
          item.child(h("div", "label").html(it.label));
        }
        return item;
      }
    });
    this.filterItems = items;
    if (items.length <= 0) {
      return;
    }
    const { el } = this;
    // items[0].toggle();
    el.html("")
      .children(...items)
      .show();
    bindClickoutside(el.parent(), () => {
      this.hide();
    });
  }

  bindInputEvents(input) {
    input.on("keydown", (evt) => inputKeydownHandler.call(this, evt));
  }
}
