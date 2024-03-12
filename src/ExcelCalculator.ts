import * as ExcelJS from 'exceljs';
import { EXCEL_TO_MATHJS_FORMULAS } from './common/constants';
import { ICell, IWorksheet } from './types';
import { MathJsInstance, all, create } from 'mathjs';
import { buildWorksheet, convertIfToTernary } from './common/helpers';

export class ExcelCalculator {
  public filePath: string;
  public sheetName: string;
  private worksheet!: IWorksheet;
  private readonly math: MathJsInstance;
  private readonly workbook: ExcelJS.Workbook;

  constructor(filePath: string, sheetName: string) {
    this.workbook = new ExcelJS.Workbook();
    this.filePath = filePath;
    this.sheetName = sheetName;
    this.math = create(all);
  }

  public setWorksheet(worksheet: IWorksheet): void {
    this.worksheet = worksheet;
  }

  public async init(filePath?: string, sheetName?: string): Promise<void> {
    await this.workbook.xlsx.readFile(filePath ?? this.filePath);
    const excelWorksheet = this.workbook.getWorksheet(sheetName ?? this.sheetName);
    if (!excelWorksheet) {
      throw new Error(`Worksheet ${sheetName} not found`);
    }
    this.worksheet = buildWorksheet(excelWorksheet);
  }

  public calculate(maxIterations: number = 10): IWorksheet {
    this.validateInitialised();
    let notCalculatedCells = this.getAllNotCalculated();
    let iterationCount = 0;
    while (notCalculatedCells.length && iterationCount < maxIterations) {
      this.calculateOnce();
      notCalculatedCells = this.getAllNotCalculated();
      iterationCount += 1;
    }
    if (iterationCount >= maxIterations) {
      console.warn('Max calculation iterations reached, there might be unresolved formulas or circular dependencies.');
    }
    return this.worksheet;
  }

  public setCellsValues(cells: Record<string, number | string | any>): void {
    this.validateInitialised();
    Object.keys(cells).forEach((cellAddress) => {
      this.setCellValue(cellAddress, cells[cellAddress]);
    });
  }

  public setCellFormula(cellAddress: string, formula: string): void {
    this.validateInitialised();
    if (!this.worksheet[cellAddress]) {
      this.worksheet[cellAddress] = { formula, value: '' };
    }
    this.worksheet[cellAddress].formula = formula;
  }

  public setCellValue(cellAddress: string, value: number | string | any): void {
    this.validateInitialised();
    if (!this.worksheet[cellAddress]) {
      this.worksheet[cellAddress] = { value };
    }
    this.worksheet[cellAddress].value = value;
  }

  public getCellFormula(cellAddress: string): string | null {
    this.validateInitialised();
    const cell = this.worksheet[cellAddress];
    return cell?.formula ? cell.formula : null;
  }

  public getCellValue(cellAddress: string): number | string | null {
    this.validateInitialised();
    const cell = this.worksheet[cellAddress.toUpperCase()];
    // eslint-disable-next-line @typescript-eslint/prefer-optional-chain
    return cell && cell.value !== undefined ? cell.value : null;
  }

  public getWorksheet(): IWorksheet {
    this.validateInitialised();
    return this.worksheet;
  }

  private validateInitialised(): void {
    if (!this.worksheet) {
      throw new Error('Worksheet not initialized. Call init() or setWorksheet() first');
    }
  }

  private replaceCellRefsWithValues(formula: string): string {
    const formulaWithValues = formula.replace(/([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)/gu, (_, col, row, col2, row2) => {
      let result = '';
      for (let i = parseInt(row, 10); i <= parseInt(row2, 10); i += 1) {
        result += `${col}${i},`;
      }
      return result.slice(0, -1);
    });
    return formulaWithValues.replace(/\$?([A-Z]+)\$?([0-9]+)/gu, (_, col, row) => {
      const cellValue = this.getCellValue(`${col}${row}`);
      return cellValue === null ? 'undefined' : cellValue.toString();
    });
  }

  private evaluateFormula(formula: string): number | string | null {
    const formulaParsed = convertIfToTernary(formula);
    let formulaWithValues = this.replaceCellRefsWithValues(formulaParsed);
    Object.entries(EXCEL_TO_MATHJS_FORMULAS).forEach(([excelFunction, mathjsFunction]) => {
      formulaWithValues = formulaWithValues.replace(new RegExp(`${excelFunction}\\(`, 'gu'), `${mathjsFunction}(`);
    });
    formulaWithValues = formulaWithValues.replace(/undefined/gu, '0');
    formulaWithValues = formulaWithValues.replace(/(?<!<)(?<!>)(?<![:])(=)(?!=)/gu, '=$1');
    try {
      return this.math.evaluate(formulaWithValues);
    } catch (error) {
      console.error('Error evaluating formula:', formula, error);
      return null;
    }
  }

  private calculateOnce(): IWorksheet {
    Object.keys(this.worksheet).forEach((key) => {
      const cell = this.worksheet[key];
      if (cell.formula) {
        const evaluatedValue = this.evaluateFormula(cell.formula);
        // eslint-disable-next-line no-eq-null
        if (evaluatedValue != null) {
          cell.value = evaluatedValue;
          delete cell.formula;
        }
      }
    });
    return this.worksheet;
  }

  private getAllNotCalculated(): ICell[] {
    return Object.values(this.worksheet).filter((cell) => cell.formula);
  }
}
