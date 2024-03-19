/* eslint-disable max-lines */
import * as ExcelJS from 'exceljs';
import { EXCEL_TO_MATHJS_FORMULAS } from './common/constants';
import { Graph, alg } from 'graphlib';
import { IWorksheet } from './types';
import { MathJsInstance, all, create } from 'mathjs';
import { buildWorksheet, convertIfToTernary, customINDEX, customMATCH, customVLOOKUP } from './common/helpers';

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
        INDEX: customINDEX,
        MATCH: customMATCH,
        VLOOKUP: customVLOOKUP,
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

    const graph = this.createDependencyGraph();
    // remove cycles
    const dagGraph = this.removeCycles(graph);
    const sortedCells = this.performTopologicalSort(dagGraph);
    this.calculateFormulas(sortedCells);

    return this.worksheet;
  }

  private createDependencyGraph(): Graph {
    const graph = new Graph();
    for (const cellAddress in this.worksheet) {
      const cell = this.worksheet[cellAddress];
      if (cell.formula) {
        graph.setNode(cellAddress);
        const dependencies = this.extractDependencies(cell.formula);
        dependencies.forEach((dependency) => {
          graph.setEdge(dependency, cellAddress);
        });
      }
    }
    return graph;
  }

  private performTopologicalSort(graph: Graph): string[] {
    try {
      return alg.topsort(graph);
    } catch (e) {
      throw new Error('Circular dependency detected');
    }
  }

  private removeCycles(graph: Graph): Graph {
    // Check if the graph is acyclic
    if (!alg.isAcyclic(graph)) {
      // Find all cycles in the graph
      const cycles = alg.findCycles(graph);
      cycles.forEach((cycle) => {
        // Remove an edge from each cycle to break the cycle
        for (let i = 0; i < cycle.length - 1; i++) {
          if (graph.hasEdge(cycle[i], cycle[i + 1])) {
            graph.removeEdge(cycle[i], cycle[i + 1]);
            break;
          }
        }
      });
    }
    return graph;
  }

  private calculateFormulas(sortedCells: string[]): void {
    sortedCells.forEach((cellAddress) => {
      const cell = this.worksheet[cellAddress];
      if (cell?.formula) {
        const evaluatedValue = this.evaluateFormula(cellAddress, cell.formula);
        if (evaluatedValue != null) {
          cell.value = evaluatedValue;
          delete cell.formula;
        }
      }
    });
  }

  private extractDependencies(formula: string): string[] {
    const regex = /\$?([A-Z]+)\$?([0-9]+)(:\$?([A-Z]+)\$?([0-9]+))?/gu;
    const matches = formula.match(regex);
    const dependencies: string[] = [];

    matches?.forEach((match) => {
      const [startCell, endCell] = match.split(':');
      if (endCell) {
        const [startCol, startRow] = this.splitCellAddress(startCell);
        const [endCol, endRow] = this.splitCellAddress(endCell);
        for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
          for (let row = startRow; row <= endRow; row++) {
            dependencies.push(String.fromCharCode(col) + row);
          }
        }
      } else {
        dependencies.push(startCell);
      }
    });
    return dependencies.map((dependency) => dependency.replace(/\$/gu, ''));
  }

  private splitCellAddress(cellAddress: string): [string, number] {
    const col = cellAddress.match(/[A-Z]+/)?.[0] ?? '';
    const row = parseInt(cellAddress.match(/[0-9]+/)?.[0] ?? '', 10);
    return [col, row];
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
    // Handle range references (e.g., A1:A3)
    let replacedFormula = formula.replace(/([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)/gu, (match, startCol, startRow, endCol, endRow) => {
      const startRowNum = parseInt(startRow, 10);
      const endRowNum = parseInt(endRow, 10);
      const rangeValues = this.getRangeValues(startCol, startRowNum, endCol, endRowNum);
      const str = JSON.stringify(rangeValues);
      return str;
    });

    // Handle normal cell references (e.g., A1, B2)
    replacedFormula = replacedFormula.replace(/\$?([A-Z]+)\$?([0-9]+)/gu, (_, col, row) => {
      const cellValue = this.getCellValue(`${col}${row}`);
      return cellValue === null ? 'undefined' : JSON.stringify(cellValue);
    });

    return replacedFormula;
  }

  private getRangeValues(startCol: string, startRow: number, endCol: string, endRow: number): any[][] {
    const rangeValues: any[][] = [];
    for (let row = startRow; row <= endRow; row++) {
      const rowValues: any[] = [];
      for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
        const cellValue = this.getCellValue(String.fromCharCode(col) + row);
        rowValues.push(cellValue === null ? 'undefined' : cellValue);
      }
      rangeValues.push(rowValues);
    }
    return rangeValues;
  }

  private evaluateFormula(cell: string, formula: string): number | string | null {
    const formulaParsed = convertIfToTernary(formula);
    let formulaWithValues = this.replaceCellRefsWithValues(formulaParsed);
    Object.entries(EXCEL_TO_MATHJS_FORMULAS).forEach(([excelFunction, mathjsFunction]) => {
      formulaWithValues = formulaWithValues.replace(new RegExp(`${excelFunction}\\(`, 'gu'), `${mathjsFunction}(`);
    });
    formulaWithValues = formulaWithValues.replace(/(?<!<)(?<!>)(?<![:])(=)(?!=)/gu, '=$1').replace(/undefined/gu, '0');
    try {
      // console.log('Evaluating cell', cellAddress, 'formula:', formulaWithValues);
      const result = this.math.evaluate(formulaWithValues);
      if (result === -0) {
        return 0;
      }
      return result;
    } catch (error) {
      console.error('Error evaluating formula:', formula, error);
      return null;
    }
  }
}
