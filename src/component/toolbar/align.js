import DropdownItem from "./dropdown_item";
import DropdownAlign from "../dropdown_align";
import { ALIGN_ITEMS } from "../../constants";

export default class Align extends DropdownItem {
  constructor(value) {
    super("align", "", value);
  }

  dropdown() {
    const { value } = this;
    return new DropdownAlign(ALIGN_ITEMS, value);
  }
}
