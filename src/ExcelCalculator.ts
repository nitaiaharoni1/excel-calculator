import * as ExcelJS from 'exceljs';
import { EXCEL_TO_MATHJS_FORMULAS } from './common/constants';
import { all, create } from 'mathjs';

const math = create(all);

export interface ICell {
  value?: number | string | any;
  formula?: string;
}

export type IWorksheet = Record<string, ICell>;

export class ExcelCalculator {
  private worksheet!: IWorksheet;
  private readonly workbook: ExcelJS.Workbook;
  public filePath: string;
  public sheetName: string;

  constructor(filePath: string, sheetName: string) {
    this.workbook = new ExcelJS.Workbook();
    this.filePath = filePath;
    this.sheetName = sheetName;
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
    this.worksheet = this.buildWorksheet(excelWorksheet);
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
    return this.cellValue(cellAddress);
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

  private cellValue(ref: string): number | string | null {
    const cell = this.worksheet[ref.toUpperCase()];
    // eslint-disable-next-line @typescript-eslint/prefer-optional-chain
    return cell && cell.value !== undefined ? cell.value : null;
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
      const cellValue = this.cellValue(`${col}${row}`);
      return cellValue === null ? 'undefined' : cellValue.toString();
    });
  }

  private parseIfContent(ifContent: string): string[] {
    const segments: string[] = [];
    let currentSegment = '';
    let openParens = 0;
    let inQuotes = false;

    for (let i = 0; i < ifContent.length; i += 1) {
      const char = ifContent[i];
      if (char === '"' && (i === 0 || ifContent[i - 1] !== '\\')) {
        // Toggle inQuotes state, ignoring escaped quotes
        inQuotes = !inQuotes;
      } else if (!inQuotes) {
        if (char === '(') {
          openParens += 1;
        } else if (char === ')') {
          openParens -= 1;
        } else if (char === ',' && openParens === 0) {
          // Only consider a comma as a segment delimiter if not within parentheses or quotes
          segments.push(currentSegment.trim());
          currentSegment = '';
          continue;
        }
      }
      currentSegment += char;
    }

    segments.push(currentSegment.trim()); // Add the last segment
    return segments;
  }

  private convertIfToTernary(formula: string): string {
    const startIndex = formula.indexOf('IF(');
    if (startIndex === -1) {
      return formula;
    }

    let endIndex = startIndex + 3;
    let openParens = 1;
    while (endIndex < formula.length && openParens > 0) {
      if (formula[endIndex] === '(') {
        openParens += 1;
      } else if (formula[endIndex] === ')') {
        openParens -= 1;
      }
      endIndex += 1;
    }

    const ifContent = formula.substring(startIndex + 3, endIndex - 1);
    const parts = this.parseIfContent(ifContent);

    if (parts.length === 3) {
      const ternaryExpression = `${parts[0]} ? ${parts[1]} : ${parts[2]}`;
      return formula.substring(0, startIndex) + ternaryExpression + formula.substring(endIndex);
    }
    return formula; // In case of unexpected format
  }

  private evaluateFormula(formula: string): number | string | null {
    const formulaParsed = this.convertIfToTernary(formula);
    let formulaWithValues = this.replaceCellRefsWithValues(formulaParsed);
    Object.entries(EXCEL_TO_MATHJS_FORMULAS).forEach(([excelFunction, mathjsFunction]) => {
      formulaWithValues = formulaWithValues.replace(new RegExp(`${excelFunction}\\(`, 'gu'), `${mathjsFunction}(`);
    });
    formulaWithValues = formulaWithValues.replace(/undefined/gu, '0');
    formulaWithValues = formulaWithValues.replace(/(?<!<)(?<!>)(?<![:])(=)(?!=)/gu, '=$1');
    try {
      return math.evaluate(formulaWithValues);
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

  private buildWorksheet(excelWorksheet: ExcelJS.Worksheet): IWorksheet {
    const worksheet: IWorksheet = {};
    excelWorksheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (!cell.address && !cell.result && !cell.formula) {
          return;
        }
        worksheet[cell.address] = {
          formula: cell.formula,
          value: cell.formula ? cell.result : cell.value,
        };
      });
    });
    return worksheet;
  }
}
