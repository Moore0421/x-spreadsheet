import { cssPrefix } from "../config";
import { h } from "./element";
import { bindClickoutside, unbindClickoutside } from "./event";
import { tf } from "../locale/locale";

function insertComment() {
  const { textEl, config, sheet } = this;
  const { data } = sheet;
  const { selector } = data;
  const { ri, ci } = selector;
  const cell = data.rows.getCellOrNew(ri, ci);
  const { onCommentAddClick, authorName, userId } = config;
  const value = textEl.el.value;
  if (value) {
    if (onCommentAddClick) {
      onCommentAddClick(cell, ri, ci, value);
    } else {
      const { c } = cell;
      const newObj = {
        a: authorName ?? "",
        t: value,
        T: true,
        timestamp: new Date().toISOString(),
        ...(userId && { userId }),
      };
      if (c) {
        c.push(newObj);
      } else {
        cell.c = [newObj];
      }
      handleReplyClick.call(this);
    }
  }
}

async function handleReplyClick() {
  handleCancelClick.call(this);
  await showAddedComment.call(this);
  buildCommentBox.call(this);
}

function keydownEventHandler(evt) {}

function handleCancelClick() {
  const { textEl } = this;
  textEl.val("");
}

function buildCommentBox() {
  const { el, textEl, sheet } = this;
  const { data } = sheet;
  const cell = data.getSelectedCell() ?? {};
  const { c = [] } = cell;
  const commentButton = h("button", `${cssPrefix}-comment-reply-button`)
    .child(tf(c?.length ? "comment.reply" : "comment.comment")())
    .on("click", () => insertComment.call(this));

  const cancelButton = h("button", `${cssPrefix}-comment-cancel-button`)
    .child(tf("comment.clear")())
    .on("click", () => handleCancelClick.call(this));

  const buttonContainer = h("div", `${cssPrefix}-comment-button-container`);
  const replyBox = h("div", `${cssPrefix}-comment-reply-box`);

  buttonContainer.child(cancelButton).child(commentButton);

  replyBox.child(textEl).child(buttonContainer);
  el.child(replyBox);
}

function buildUser() {
  const { config } = this;
  const { authorName } = config;
  const avatar = h("div", `${cssPrefix}-comment-avatar`).child(
    authorName ? authorName.toString()[0]?.toUpperCase() : ""
  );
  const name = h("div", `${cssPrefix}-comment-name`).child(authorName);
  const box = h("div", `${cssPrefix}-title-box`).child(name);
  const header = h("div", `${cssPrefix}-comment-inner-box`)
    .child(avatar)
    .child(box);
  const div = h("div", `${cssPrefix}-inner-box`).child(header);
  return div;
}

function buildCommentStack(comments) {
  const {
    config: { enableTimeStamp },
  } = this;
  const ele = comments.map((cmt) => {
    const { a, t, timestamp } = cmt;
    const avatar = h("div", `${cssPrefix}-comment-avatar`).child(
      a ? a.toString()[0]?.toUpperCase() : ""
    );
    const name = h("div", `${cssPrefix}-comment-name`).child(a);
    const box = h("div", `${cssPrefix}-title-box`).child(name);
    if (enableTimeStamp) {
      const dateString = new Date(timestamp).toLocaleString();
      const date = h("div", `${cssPrefix}-comment-data`).child(
        dateString ?? ""
      );
      box.child(date);
    }
    const header = h("div", `${cssPrefix}-comment-inner-box`)
      .child(avatar)
      .child(box);
    const contentBox = h("div", `${cssPrefix}-content-box`).child(t ?? "");
    const div = h("div", `${cssPrefix}-inner-box`)
      .child(header)
      .child(contentBox);
    return div;
  });
  return ele;
}

function removeComments() {
  const { el } = this;
  el.html("");
}

async function showAddedComment() {
  removeComments.call(this);
  const { el, sheet = {}, config, loader } = this;
  const { showComments } = config;
  const { data = {} } = sheet;
  const cell = data.getSelectedCell();
  if (cell) {
    el.child(loader);
    loader.show();
    let cmtArray;
    if (showComments) {
      cmtArray = await showComments?.(cell);
    } else {
      cmtArray = cell.c ?? [];
    }
    const elements = buildCommentStack.call(this, cmtArray);
    loader.hide();
    if (elements?.length) {
      el.children(...elements);
    } else {
      const ele = buildUser.call(this);
      el.child(ele);
    }
  }
}

function removeAllChildren(el) {
  el.html("");
  el.hide();
  unbindClickoutside(el);
}

export default class Comment {
  constructor(sheet = {}, viewFn, isHide = false) {
    this.sheet = sheet;
    this.config = this.sheet.options?.comment ?? {};
    this.viewFn = viewFn;
    this.data = [];
    this.isHide = isHide;
    this.loader = h("div", `${cssPrefix}-loader`).child("Loading...").hide();
    this.textEl = h("textarea", `${cssPrefix}-comment-textarea`).on(
      "keydown",
      (evt) => keydownEventHandler.call(this, evt)
    );
    this.el = h("div", `${cssPrefix}-comment-box`).hide();
  }

  hide() {
    const { el } = this;
    el.hide();
    unbindClickoutside(el);
  }

  async show() {
    if (!this?.config.indicatorColor) return;
    if (this.isHide) return;
    const { el, sheet = {} } = this;
    const { contextMenu } = sheet;
    await showAddedComment.call(this);
    buildCommentBox.call(this);
    const { x, y } = contextMenu.lastCoordinate;
    const { width } = el.show().offset();
    const view = this.viewFn();
    const vhf = view.height / 2;
    let left = x;
    if (view.width - x <= width) {
      left -= width;
    }

    el.css("left", `${left}px`);
    if (y > vhf) {
      el.css("bottom", `${view.height - y}px`)
        .css("max-height", `${y}px`)
        .css("top", "auto");
    } else {
      el.css("top", `${y}px`)
        .css("max-height", `${view.height - y}px`)
        .css("bottom", "auto");
    }
    setTimeout(() => {
      bindClickoutside(el, removeAllChildren);
    }, 100);
  }
}
