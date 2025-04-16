declare module "x-data-spreadsheet" {
  export interface ExtendToolbarOption {
    id?: string;
    tip?: string;
    el?: HTMLElement;
    icon?: string;
    onClick?: (data: object, sheet: object) => void;
  }
  export interface Options {
    mode?: "edit" | "read";
    showToolbar?: boolean;
    showGrid?: boolean;
    showContextmenu?: boolean;
    showBottomBar?: boolean;
    extendToolbar?: {
      left?: ExtendToolbarOption[];
      right?: ExtendToolbarOption[];
    };
    autoFocus?: boolean;
    view?: {
      height: () => number;
      width: () => number;
    };
    row?: {
      len: number;
      height: number;
      minHeight?: number;
    };
    col?: {
      len: number;
      width: number;
      indexWidth: number;
      minWidth?: number;
    };
    style?: {
      bgcolor: string;
      align: "left" | "center" | "justify" | "right";
      valign: "top" | "middle" | "bottom";
      textwrap: boolean;
      strike: boolean;
      underline: boolean;
      color: string;
      font: {
        name: "Helvetica";
        size: number;
        bold: boolean;
        italic: false;
      };
    };
    cellConfigButtons?: CellConfigButton[];
    numberConfig?: { zeroReplacer?: string; groupingSymbol?: string };
    valueFormatter?: (metaData: FormatterMeta) => string;
    valueSetter?: (metaData: FormatterMeta) => string;
    editValueFormatter?: (metaData: FormatterMeta) => string;
    cellStyleProvider?: (attr: object, metaData: FormatterMeta) => object;
    onFormulaCellFinalized?: (params: Record<string, any>) => void;
    extendedContextMenu?: ExtendedContextMenu[];
    mentionProgress?: {
      trigger: string;
      stepCount?: number;
      itemCall?: (word: string) => Promise<any>;
    };
    allowMultipleSheets?: boolean;
    disableFeatures?: string[];
    suppressMaximumSelection?: boolean;
    comment?: {
      userId?: string;
      authorName?: string;
      indicatorColor?: string;
      enableTimeStamp?: boolean;
      onCommentAddClick: (
        cell: any,
        ri: number,
        ci: number,
        text: string
      ) => Promise;
      showComments: (cell: any) => Promise;
    };
  }

  export type CELL_SELECTED = "cell-selected";
  export type CELLS_SELECTED = "cells-selected";
  export type CELL_EDITED = "cell-edited";
  export type TOOLBAR_ACTION = "toolbar-action";
  export type CONTEXT_MENU_ACTION = "context-menu-action";
  export type SHEET_CHANGE = "sheet-change";
  export type PASTED_CLIPBOARD = "pasted-clipboard";
  export type CUT_CLIPBOARD = " cut-clipboard";
  export type COPIED_CLIPBOARD = "copied-clipboard";
  export type GRID_LOAD = "grid-load";
  export type CELL_EDIT_FINISHED = "cell-edit-finished";
  export type ROW_COLUMN_OPERATION = "row-column-operation";

  export interface ExtendedContextMenu {
    key: string;
    title: string;
    callback?: (
      ri: number,
      ci: number,
      range: CellRangeType,
      cellMeta: CellData
    ) => void;
    subMenus: ExtendedContextMenu[];
    visibility?: (sheet: any, key: string) => boolean;
  }

  export interface FormatterMeta {
    value: string;
    cell: CellData;
  }

  export interface CellConfigButton {
    id?: string;
    tag: string;
    tip?: string;
    icon?: string;
    indicator: string;
  }

  export interface SheetChangeType {
    action: string;
    sheet: any;
  }

  export interface ToolbarOrContextActionType {
    action: any[];
    range: CellRangeType;
  }

  export type CellMerge = [number, number];

  export interface SpreadsheetEventHandler {
    (
      envt: CELL_SELECTED,
      callback: (cell: Cell, rowIndex: number, colIndex: number) => void
    ): void;
    (
      envt: CELLS_SELECTED,
      callback: (
        cell: Cell,
        parameters: { sri: number; sci: number; eri: number; eci: number }
      ) => void
    ): void;
    (
      evnt: CELL_EDITED,
      callback: (text: string, rowIndex: number, colIndex: number) => void
    ): void;
    (
      evnt: TOOLBAR_ACTION,
      callback: (data: ToolbarOrContextActionType) => void
    ): void;
    (
      evnt: CONTEXT_MENU_ACTION,
      callback: (data: ToolbarOrContextActionType) => void
    ): void;
    (evnt: SHEET_CHANGE, callback: (data: SheetChangeType) => void): void;
    (evnt: PASTED_CLIPBOARD, callback: (data: any) => void): void;
    (evnt: CUT_CLIPBOARD, callback: (data: any) => void): void;
    (evnt: COPY_CLIPBOARD, callback: (data: any) => void): void;
    (evnt: GRID_LOAD, callback: (data: any) => void): void;
    (
      evnt: ROW_COLUMN_OPERATION,
      callback: (data: RowColumnOperationProp) => void
    ): void;
    (
      evnt: CELL_EDIT_FINISHED,
      callback: (text: string, rowIndex: number, colIndex: number) => void
    ): void;
  }

  export interface RowColumnOperationProp {
    on: string;
    type: string;
    startIndex: number;
    count: number;
  }

  export interface CellRangeType {
    sci: number;
    sri: number;
    eci: number;
    eri: number;
  }

  export interface ColProperties {
    width?: number;
  }

  /**
   * Data for representing a cell
   */
  export interface CellData {
    text: string;
    style?: number;
    merge?: CellMerge;
    cellMeta?: { [key: string]: boolean };
  }
  /**
   * Data for representing a row
   */
  export interface RowData {
    cells: {
      [key: number]: CellData;
    };
  }

  /**
   * Data for representing a sheet
   */
  export interface SheetData {
    name?: string;
    freeze?: string;
    styles?: CellStyle[];
    merges?: string[];
    cols?: {
      len?: number;
      [key: number]: ColProperties;
    };
    rows?: {
      [key: number]: RowData;
    };
  }

  /**
   * Data for representing a spreadsheet
   */
  export interface SpreadsheetData {
    [index: number]: SheetData;
  }

  export interface CellStyle {
    align?: "left" | "center" | "justify" | "right";
    valign?: "top" | "middle" | "bottom";
    font?: {
      bold?: boolean;
    };
    bgcolor?: string;
    textwrap?: boolean;
    color?: string;
    border?: {
      top?: string[];
      right?: string[];
      bottom?: string[];
      left?: string[];
    };
  }
  export interface Editor {}
  export interface Element {}

  export interface Row {}
  export interface Table {}
  export interface Cell {}
  export interface Sheet {}

  export default class Spreadsheet {
    constructor(container: string | HTMLElement, opts?: Options);
    on: SpreadsheetEventHandler;
    /**
     * retrieve cell
     * @param rowIndex {number} row index
     * @param colIndex {number} column index
     * @param sheetIndex {number} sheet iindex
     */
    cell(rowIndex: number, colIndex: number, sheetIndex: number): Cell;
    /**
     * retrieve cell style
     * @param rowIndex
     * @param colIndex
     * @param sheetIndex
     */
    cellStyle(
      rowIndex: number,
      colIndex: number,
      sheetIndex: number
    ): CellStyle;
    /**
     * get/set cell text
     * @param rowIndex
     * @param colIndex
     * @param text
     * @param sheetIndex
     */
    cellText(
      rowIndex: number,
      colIndex: number,
      text: string,
      sheetIndex?: number
    ): this;
    /**
     * remove current sheet
     */
    deleteSheet(): void;

    /**s
     * load data
     * @param json
     */
    loadData(json: Record<string, any>): this;
    /**
     * get data
     */
    getData(): Record<string, any>;
    /**
     * bind handler to change event, including data change and user actions
     * @param callback
     */
    change(callback: (json: Record<string, any>) => void): this;
    /**
     * set locale
     * @param lang
     * @param message
     */
    static locale(lang: string, message: object): void;
    /**
     * selects the cell
     * @param rowIndex
     * @param colIndex
     */
    selectCell(rowIndex: number, colIndex: number): void;
  }
  global {
    interface Window {
      x_spreadsheet(
        container: string | HTMLElement,
        opts?: Options
      ): Spreadsheet;
    }
  }
}
