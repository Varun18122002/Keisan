// spreadsheet.types.ts

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  backgroundColor?: string;
}

export interface CellData {
  value: any;
  style?: CellStyle;
}

export interface SheetData {
  name: string;
  data: CellData[][];
  id: number;
  columnHeaders: string[];
}

export interface SpreadsheetFile {
  sheets: SheetData[];
  activeSheet: number;
}

export type FileFormat = 'xlsx' | 'csv';

export interface FileDialogOptions {
  defaultPath?: string;
  filters: {
    name: string;
    extensions: string[];
  }[];
}
