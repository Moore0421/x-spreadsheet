import { cssPrefix } from "../../config";
import Button from "../button";
import { h } from "../element";
import { xtoast } from "../message";
import Modal from "../modal";
import FormField from "../form_field";
import FormInput from "../form_input";

class SheetIdModal {
  constructor() {
    this.sheetId = '';
    this.modalEl = null;
    this.callback = null;
  }

  render(defaultValue = '') {
    // 创建输入框
    const InputFormField = new FormField(
      new FormInput("200px", "请输入表ID")
        .val(defaultValue),
      {
        required: true,
      },
      "表ID:",
      100
    );

    const InputFormFieldWrapper = h("div", `${cssPrefix}-form-fields`).children(
      InputFormField.el
    );

    const Buttons = h("div", `${cssPrefix}-buttons`).children(
      new Button("cancel").on("click", () => {
        this.sheetId = '';
        this.modalEl.hide();
      }).el,
      new Button("ok", "primary").on("click", () => {
        const value = InputFormField.input.val();
        if (!value) {
          xtoast("错误提示", "请输入表ID");
          return;
        }
        this.handleConfirm(value);
        this.modalEl.hide();
      }).el
    );

    const modal = new Modal(
      "设置表ID",
      [InputFormFieldWrapper.el, Buttons.el],
      "400px"
    );

    this.modalEl = modal;
    return modal;
  }

  handleConfirm(sheetId) {
    if (this.callback) {
      this.callback(sheetId);
    }
  }

  show(callback, defaultValue = '') {
    this.callback = callback;
    this.modalEl = this.render(defaultValue);
    this.modalEl.show();
  }
}

export default SheetIdModal;