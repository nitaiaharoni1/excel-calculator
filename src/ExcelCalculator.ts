/* eslint-disable max-lines */
import * as ExcelJS from 'exceljs';
import { EXCEL_TO_MATHJS_FORMULAS, MULTIPLE_ARGS_FORMULAS } from './common/constants';
import { IWorksheet } from './types';
import { MathJsInstance, all, create } from 'mathjs';
import { buildWorksheet, convertIfToTernary, customIndexFunction, customMatchFunction, customVlookupFunction } from './common/helpers';

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

    this.math.import(
      {
        INDEX: customIndexFunction,
        MATCH: customMatchFunction,
        VLOOKUP: customVlookupFunction,
      },
      { override: true },
    );
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

  public calculate(): IWorksheet {
    this.validateInit();
    let iteration = 0;
    while (iteration < 10) {
      let formulaFound = false;
      for (const cellAddress in this.worksheet) {
        const cell = this.worksheet[cellAddress];
        if (cell.formula && cell.formula !== '#REF!') {
          // Skip evaluation for cells marked with #REF!
          const evaluatedValue = this.evaluateFormula(cell.formula);
          if (evaluatedValue != null) {
            cell.value = evaluatedValue;
            delete cell.formula;
            formulaFound = true;
          }
        }
      }
      if (!formulaFound) {
        break;
      }
      iteration += 1;
    }
    return this.worksheet;
  }

  public setCellsValues(cells: Record<string, number | string | any>): void {
    this.validateInit();
    Object.keys(cells).forEach((cellAddress) => {
      this.setCellValue(cellAddress, cells[cellAddress]);
    });
  }

  public setCellFormula(cellAddress: string, formula: string): void {
    this.validateInit();
    if (!this.worksheet[cellAddress]) {
      this.worksheet[cellAddress] = { formula, value: '' };
    }
    this.worksheet[cellAddress].formula = formula;
  }

  public setCellValue(cellAddress: string, value: number | string | any): void {
    this.validateInit();
    if (!this.worksheet[cellAddress]) {
      this.worksheet[cellAddress] = { value };
    }
    this.worksheet[cellAddress].value = value;
  }

  public getCellFormula(cellAddress: string): string | null {
    this.validateInit();
    const cell = this.worksheet[cellAddress];
    return cell?.formula ? cell.formula : null;
  }

  public getCellValue(cellAddress: string): number | string | null {
    this.validateInit();
    const cell = this.worksheet[cellAddress.toUpperCase()];
    // eslint-disable-next-line @typescript-eslint/prefer-optional-chain
    return cell && cell.value !== undefined ? cell.value : null;
  }

  public getWorksheet(): IWorksheet {
    this.validateInit();
    return this.worksheet;
  }

  private validateInit(): void {
    if (!this.worksheet) {
      throw new Error('Worksheet not initialized. Call init() or setWorksheet() first');
    }
  }

  private replaceCellRefsWithValues(formula: string): string {
    // Now, handle range references (e.g., A1:A3). This is where we adjust for INDEX and MATCH support.
    // This regex identifies range references and replaces them with a structured array representation.
    let replacedFormula = formula;
    replacedFormula = replacedFormula.replace(/([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)/gu, (match, startCol, startRow, endCol, endRow) => {
      // Convert start and end rows to numbers for iteration

      const startRowNum = parseInt(startRow, 10);
      const endRowNum = parseInt(endRow, 10);
      const rangeValues = this.getRangeValues(startCol, startRowNum, endCol, endRowNum);
      if (MULTIPLE_ARGS_FORMULAS.some((f) => formula.includes(f))) {
        return JSON.stringify(rangeValues);
      }
      return rangeValues.flat().join(',');
    });

    // Handle normal cell references (e.g., A1, B2) as your existing logic might already do.
    // This regex identifies individual cell references and replaces them with their corresponding values.
    replacedFormula = replacedFormula.replace(/\$?([A-Z]+)\$?([0-9]+)/gu, (_, col, row) => {
      const cellValue = this.getCellValue(`${col}${row}`);
      return cellValue === null ? 'undefined' : JSON.stringify(cellValue); // Convert to JSON string to handle strings correctly in formulas
    });

    return replacedFormula;
  }

  private getRangeValues(startCol: string, startRow: number, endCol: string, endRow: number): any[][] {
    const rangeValues = [];
    // Assume column letters can be converted to and from ASCII codes for simplicity.
    // For a more robust solution, you might need a method to convert column letters to numbers and vice versa.
    for (let rowNum = startRow; rowNum <= endRow; rowNum += 1) {
      const rowValues = [];
      for (let colCode = startCol.charCodeAt(0); colCode <= endCol.charCodeAt(0); colCode += 1) {
        const colLetter = String.fromCharCode(colCode);
        const cellValue = this.getCellValue(`${colLetter}${rowNum}`);
        rowValues.push(cellValue === null ? undefined : cellValue); // Use undefined for empty cells to mirror Excel behavior
      }
      rangeValues.push(rowValues);
    }
    return rangeValues;
  }

  private evaluateFormula(formula: string): number | string | null {
    const formulaParsed = convertIfToTernary(formula);
    let formulaWithValues = this.replaceCellRefsWithValues(formulaParsed);
    if (formulaWithValues.includes('undefined')) {
      return null;
    }
    Object.entries(EXCEL_TO_MATHJS_FORMULAS).forEach(([excelFunction, mathjsFunction]) => {
      formulaWithValues = formulaWithValues.replace(new RegExp(`${excelFunction}\\(`, 'gu'), `${mathjsFunction}(`);
    });
    // formulaWithValues = formulaWithValues.replace(/undefined/gu, '0');
    formulaWithValues = formulaWithValues.replace(/(?<!<)(?<!>)(?<![:])(=)(?!=)/gu, '=$1');
    try {
      return this.math.evaluate(formulaWithValues);
    } catch (error) {
      console.error('Error evaluating formula:', formula, error);
      return null;
    }
  }
}
