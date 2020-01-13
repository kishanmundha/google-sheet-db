export interface ISheetProperty {
  gridProperties?: {
    columnCount?: number;
    columnGroupControlAfter?: boolean;
    frozenColumnCount?: number;
    frozenRowCount?: number;
    hideGridlines?: boolean;
    rowCount?: number;
    rowGroupControlAfter?: boolean;
  };
  hidden?: boolean;
  index?: number;
  rightToLeft?: boolean;
  sheetId?: number;
  sheetType?: string;
  tabColor?: {
    alpha?: number;
    blue?: number;
    green?: number;
    red?: number;
  };
  title?: string;
}
