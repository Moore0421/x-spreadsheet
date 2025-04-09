const REF_ERROR = "REF!";
const GENERAL_ERROR = "ERROR!";
const VALUE_ERROR = "VALUE!";
const DYNAMIC_VARIABLE_ERROR = "DYNAMIC_VARIABLE_ERROR!";
const CIRCULAR_DEPENDENCY_ERROR = "CIRCULAR_DEPENDENCY_ERROR!";
const DYNAMIC_VARIABLE_RESOLVING = "RESOLVING_VARIABLE!";

const CELL_REF_REGEX = /\b[A-Za-z]+[1-9][0-9]*\b/g;
const CELL_RANGE_REGEX =
  /\$?[A-Za-z]+\$?[1-9][0-9]*:\$?[A-Za-z]+\$?[1-9][0-9]*/gi;
const SPACE_REMOVAL_REGEX = /\s+(?=(?:[^']*'[^']*')*[^']*$)/g;
const SHEET_TO_CELL_REF_REGEX =
  /(?:'([^']*)'|\b[A-Za-z0-9]+)\![A-Za-z]+[1-9][0-9]*/g;
const CELL_REF_REPLACE_REGEX = /(?:\b[A-Za-z0-9_]+\b!)?[A-Za-z]+\d+\b/g;
const ACCOUNTING_FORMAT_REGEX = /^\s*[€£¥₹$¢₱₦₩]?[\s]*\(?-?[\d,]+(\.\d+)?\)?$/;
const EXTRACT_FORMULA_CELL_NAME_REGEX = /\b[a-z]+\b(?=\()|(?<=\W)[a-z]\d+\b/gi;
const EXTRACT_MSO_NUMBER_FORMAT_REGEX =
  /\.([^ \{]+)\s*{[^}]*mso-number-format:\s*"([^"]+)"/g;

const PX_TO_PT = 0.75;

const AVAILABLE_FEATURES = [
  "import",
  "undo",
  "redo",
  "print",
  "paintFormat",
  "clearFormat",
  "format",
  "font",
  "fontSize",
  "bold",
  "italic",
  "underline",
  "strike",
  "textColor",
  "fillColor",
  "border",
  "merge",
  "grid",
  "align",
  "vAlign",
  "textWrap",
  "freeze",
  "autoFilter",
  "formula",
];

const INDEXED_COLORS = {
  0: "#000000",
  1: "#FFFFFF",
  2: "#FF0000",
  3: "#00FF00",
  4: "#0000FF",
  5: "#FFFF00",
  6: "#FF00FF",
  7: "#00FFFF",
  8: "#000000",
  9: "#FFFFFF",
  10: "#FF0000",
  11: "#00FF00",
  12: "#0000FF",
  13: "#FFFF00",
  14: "#FF00FF",
  15: "#00FFFF",
  16: "#800000",
  17: "#008000",
  18: "#000080",
  19: "#808000",
  20: "#800080",
  21: "#008080",
  22: "#C0C0C0",
  23: "#808080",
  24: "#9999FF",
  25: "#993366",
  26: "#FFFFCC",
  27: "#CCFFFF",
  28: "#660066",
  29: "#FF8080",
  30: "#0066CC",
  31: "#CCCCFF",
  32: "#000080",
  33: "#FF00FF",
  34: "#FFFF00",
  35: "#00FFFF",
  36: "#800080",
  37: "#800000",
  38: "#008080",
  39: "#0000FF",
  40: "#00CCFF",
  41: "#CCFFFF",
  42: "#CCFFCC",
  43: "#FFFF99",
  44: "#99CCFF",
  45: "#FF99CC",
  46: "#CC99FF",
  47: "#FFCC99",
  48: "#3366FF",
  49: "#33CCCC",
  50: "#99CC00",
  51: "#FFCC00",
  52: "#FF9900",
  53: "#FF6600",
  54: "#666699",
  55: "#969696",
  56: "#003366",
  57: "#339966",
  58: "#003300",
  59: "#333300",
  60: "#993300",
  61: "#993366",
  62: "#333399",
  63: "#333333",
  64: "#000000",
};

const DEFAULT_ROW_HEIGHT = 20;

const EXCEL_ERRORS = [REF_ERROR, GENERAL_ERROR, VALUE_ERROR];
const DEFAULT_FORMATS = ["general", "text", "number"];

const ALIGN_ITEMS = ["left", "center", "justify", "right"];

const GENERAL_CELL_OBJECT = {
  t: "n",
  v: "",
  h: "",
  s: "",
  f: "",
  z: "General",
};

export {
  REF_ERROR,
  GENERAL_ERROR,
  CELL_REF_REGEX,
  CELL_RANGE_REGEX,
  SPACE_REMOVAL_REGEX,
  SHEET_TO_CELL_REF_REGEX,
  CELL_REF_REPLACE_REGEX,
  PX_TO_PT,
  AVAILABLE_FEATURES,
  INDEXED_COLORS,
  DEFAULT_ROW_HEIGHT,
  DYNAMIC_VARIABLE_ERROR,
  CIRCULAR_DEPENDENCY_ERROR,
  DYNAMIC_VARIABLE_RESOLVING,
  ACCOUNTING_FORMAT_REGEX,
  VALUE_ERROR,
  EXCEL_ERRORS,
  EXTRACT_FORMULA_CELL_NAME_REGEX,
  EXTRACT_MSO_NUMBER_FORMAT_REGEX,
  ALIGN_ITEMS,
  DEFAULT_FORMATS,
  GENERAL_CELL_OBJECT,
};
