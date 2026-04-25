import { cssPrefix } from "../../config";
import Button from "../button";
import { h } from "../element";
import { xtoast } from "../message";
import Modal from "../modal";
import FormField from "../form_field";
import FormInput from "../form_input";

class CellIdModal {
  constructor() {
    this.cellId = '';
    this.modalEl = null;
    this.callback = null;
  }

  render(defaultValue = '') {
    // 创建输入框
    const InputFormField = new FormField(
      new FormInput("200px", "请输入单元格ID")
        .val(defaultValue),
      {
        required: true,
      },
      "单元格ID:",
      100
    );

    const InputFormFieldWrapper = h("div", `${cssPrefix}-form-fields`).children(
      InputFormField.el
    );

    const Buttons = h("div", `${cssPrefix}-buttons`).children(
      new Button("cancel").on("click", () => {
        this.cellId = '';
        this.modalEl.hide();
      }).el,
      new Button("ok", "primary").on("click", () => {
        const value = InputFormField.input.val();
        if (!value) {
          xtoast("错误提示", "请输入单元格ID");
          return;
        }
        this.handleConfirm(value);
        this.modalEl.hide();
      }).el
    );

    const modal = new Modal(
      "设置单元格ID",
      [InputFormFieldWrapper.el, Buttons.el],
      "400px"
    );

    this.modalEl = modal;
    return modal;
  }

  handleConfirm(cellId) {
    if (this.callback) {
      this.callback(cellId);
    }
  }

  show(callback, defaultValue = '') {
    this.callback = callback;
    this.modalEl = this.render(defaultValue);
    this.modalEl.show();
  }
}

export default CellIdModal;