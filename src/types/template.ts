/**
 * Template and Placeholder Data Types
 */

export interface CellOccurrence {
  sheetName: string;
  cellAddress: string;
  row: number;
  col: number;
  originalText: string;
}

export interface PlaceholderVariable {
  key: string;
  defaultValue: string;
  currentValue: string;
  isLongText: boolean;
  isDate?: boolean;
  dropdownOptions?: string[];
  occurrences: CellOccurrence[];
  category: string;
}

export interface TemplateMeta {
  fileName: string;
  fileSize: number;
  sheetNames: string[];
  totalCellsScanned: number;
  formulaCount: number;
  imageCount: number;
  hasHeaderFooter: boolean;
  rawBuffer: ArrayBuffer;
}

export interface QuotationItem {
  id: string;
  rowNumber: number;
  nameKey: string;
  descKey: string;
  qtyKey: string;
  unitKey: string;
  priceKey: string;
  name: string;
  desc: string;
  qty: string;
  unit: string;
  price: string;
}

export interface ToastNotification {
  id: string;
  type: 'success' | 'info' | 'warning' | 'error';
  title: string;
  message: string;
}
