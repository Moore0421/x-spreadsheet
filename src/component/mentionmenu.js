import { h } from "./element";
import { bindClickoutside, unbindClickoutside } from "./event";
import { cssPrefix } from "../config";

function inputMovePrev(evt) {
  evt.preventDefault();
  evt.stopPropagation();
  const { filterItems } = this;
  if (filterItems.length <= 0) return;
  if (this.itemIndex >= 0) filterItems[this.itemIndex].toggle();
  this.itemIndex -= 1;
  if (this.itemIndex < 0) {
    this.itemIndex = filterItems.length - 1;
  }
  filterItems[this.itemIndex].toggle();
}

function inputMoveNext(evt) {
  evt.preventDefault();
  evt.stopPropagation();
  const { filterItems } = this;
  if (filterItems.length <= 0) return;
  if (this.itemIndex >= 0) filterItems[this.itemIndex].toggle();
  this.itemIndex += 1;
  if (this.itemIndex > filterItems.length - 1) {
    this.itemIndex = 0;
  }
  filterItems[this.itemIndex].toggle();
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
    case 39: // right
      evt.stopPropagation();
      break;
    case 38: // up
      inputMovePrev.call(this, evt);
      break;
    case 40: // down
      inputMoveNext.call(this, evt);
      break;
    case 13: // enter
    case 9:
      inputEnter.call(this, evt);
      break;
    default:
      evt.stopPropagation();
      break;
  }
}

export default class MentionMenu {
  constructor(menuConfig = {}, itemClick, width = "200px") {
    this.filterItems = [];
    this.items = [];
    this.trigger = menuConfig.trigger;
    this.el = h("div", `${cssPrefix}-suggest`).css("width", width).hide();
    this.itemIndex = -1;
    this.itemClick = itemClick;
    this.getItemCall = menuConfig.itemCall;
    this.loader = h("div", `${cssPrefix}-loader`).child("Loading...").hide();
    this.lastFetchedString = "";
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
  }

  async search(word) {
    if (!this.trigger.length === 1) return;
    this.lastFetchedString = word;
    const lastChar = word?.[word?.length - 1];
    let fetchCall = false;
    if (lastChar === this.trigger || lastChar === ".") {
      this.el.html("").child(this.loader).show();
      this.showLoader();
      this.items = [];
      fetchCall = true;
    }
    try {
      let items = fetchCall
        ? ((await this.getItemCall?.(word)) ?? [])
        : this.items;
      this.items = items;
      const wordSplit = word.split(".");
      const lastWord = wordSplit[wordSplit?.length - 1]?.replace(
        this.trigger,
        ""
      );
      items = items.filter((item) =>
        item.value.toLowerCase().includes(lastWord?.toLowerCase())
      );
      items = items.map((it) => {
        let { value } = it;
        const item = h("div", `${cssPrefix}-item`)
          .child(value)
          .on("click.stop", () => {
            this.itemClick(it);
            it.callback?.call(this, it);
            this.hide();
          });
        return item;
      });
      this.filterItems = items;
      this.hideLoader();
      if (items.length <= 0) {
        this.hide();
        return;
      }
      const { el } = this;
      el.html("")
        .children(...items)
        .show();
      bindClickoutside(el.parent(), () => {
        this.hide();
      });
    } catch (error) {
      console.log("catch", error);
      this.hideLoader();
      this.hide();
    }
  }

  showLoader() {
    this.loader.show();
    this.el.show();
  }

  hideLoader() {
    this.loader.hide();
  }

  bindInputEvents(input) {
    input.on("keydown", (evt) => inputKeydownHandler.call(this, evt));
  }
}
