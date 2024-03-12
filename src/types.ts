export interface ICell {
  value?: number | string | any;
  formula?: string;
}

export type IWorksheet = Record<string, ICell>;
