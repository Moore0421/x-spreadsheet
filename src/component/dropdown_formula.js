import Dropdown from "./dropdown";
import Icon from "./icon";
import { h } from "./element";
import { baseFormulas } from "../core/formula";
import { cssPrefix } from "../config";

export default class DropdownFormula extends Dropdown {
  constructor() {
    const nformulas = baseFormulas.map((it) =>
      h("div", `${cssPrefix}-item`)
        .css("height", "unset")
        .css("overflow-wrap", "break-word")
        .css("border-bottom", "0.1px solid #EBEBEB")
        .on("click", () => {
          this.hide();
          this.change(it);
        })
        .child(it.key)
    );
    super(new Icon("formula"), "200px", true, "bottom-left", ...nformulas);
  }
}
