import { ExcelCalculator } from './ExcelCalculator';

describe('ExcelCalculator Tests', () => {
  let excelCalculator: ExcelCalculator;

  beforeEach(() => {
    excelCalculator = new ExcelCalculator('dummy/path.xlsx', 'Sheet1');
    excelCalculator.setWorksheet({});
  });

  it('should calculate formulas correctly', () => {
    const mockWorksheet = {
      A1: { value: 2 },
      B1: { value: 3 },
      C1: { formula: 'A1+B1' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.C1.value).toBe(5);
  });

  it('sets cell values correctly', () => {
    const cellAddress = 'D1';
    const value = 10;
    excelCalculator.setCellValue(cellAddress, value);
    expect(excelCalculator.getCellValue('D1')).toEqual(value);
  });

  it('sets cell formula correctly', () => {
    const cellAddress = 'E1';
    const formula = 'A1+D1';
    excelCalculator.setCellFormula(cellAddress, formula);
    expect(excelCalculator.getCellFormula(cellAddress)).toEqual(formula);
  });
});
