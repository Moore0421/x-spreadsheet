const REF_ERROR = "#REF!";
const GENERAL_ERROR = "#ERROR";

const CELL_REF_REGEX = /\b[A-Za-z]+[1-9][0-9]*\b/g;
const CELL_RANGE_REGEX =
  /\$?[A-Za-z]+\$?[1-9][0-9]*:\$?[A-Za-z]+\$?[1-9][0-9]*/gi;
const SPACE_REMOVAL_REGEX = /\s+(?=(?:[^']*'[^']*')*[^']*$)/g;
const SHEET_TO_CELL_REF_REGEX =
  /(?:'([^']*)'|\b[A-Za-z0-9]+)\![A-Za-z]+[1-9][0-9]*/g;
const CELL_REF_REPLACE_REGEX = /(?:\b[A-Za-z0-9_]+\b!)?[A-Za-z]+\d+\b/g;

export {
  REF_ERROR,
  GENERAL_ERROR,
  CELL_REF_REGEX,
  CELL_RANGE_REGEX,
  SPACE_REMOVAL_REGEX,
  SHEET_TO_CELL_REF_REGEX,
  CELL_REF_REPLACE_REGEX,
};
