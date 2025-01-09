import { cssPrefix } from "../../config";
import Button from "../button";
import { h } from "../element";
import { xtoast } from "../message";
import Modal from "../modal";
import FormField from "../form_field";
import FormInput from "../form_input";

class RowIdModal {
  constructor() {
    this.rowId = '';
    this.modalEl = null;
    this.callback = null;
  }

  render(defaultValue = '') {
    // 创建输入框
    const InputFormField = new FormField(
      new FormInput("200px", "请输入行ID")
        .val(defaultValue),
      {
        required: true,
      },
      "行ID:",
      100
    );

    const InputFormFieldWrapper = h("div", `${cssPrefix}-form-fields`).children(
      InputFormField.el
    );

    const Buttons = h("div", `${cssPrefix}-buttons`).children(
      new Button("cancel").on("click", () => {
        this.rowId = '';
        this.modalEl.hide();
      }).el,
      new Button("ok", "primary").on("click", () => {
        const value = InputFormField.input.val();
        if (!value) {
          xtoast("错误提示", "请输入行ID");
          return;
        }
        this.handleConfirm(value);
        this.modalEl.hide();
      }).el
    );

    const modal = new Modal(
      "设置行ID",
      [InputFormFieldWrapper.el, Buttons.el],
      "400px"
    );

    this.modalEl = modal;
    return modal;
  }

  handleConfirm(rowId) {
    if (this.callback) {
      this.callback(rowId);
    }
  }

  show(callback, defaultValue = '') {
    this.callback = callback;
    this.modalEl = this.render(defaultValue);
    this.modalEl.show();
  }
}

export default RowIdModal;