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

  it('should throw error when worksheet is not initialized', () => {
    excelCalculator.setWorksheet(null as any);
    expect(() => excelCalculator.calculate()).toThrow('Worksheet not initialized. Call init() or setWorksheet() first');
  });

  it('should calculate complex formulas correctly', () => {
    const mockWorksheet = {
      A1: { value: 5 },
      A2: { value: 10 },
      A3: { formula: 'A1*A2+10' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A3.value).toBe(60);
  });

  it('should handle formulas referencing undefined cells', () => {
    const mockWorksheet = {
      A1: { formula: 'B1+10' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A1.value).toBe(10);
  });

  it('should handle formulas using cell ranges', () => {
    const mockWorksheet = {
      A1: { value: 1 },
      A2: { value: 2 },
      A3: { value: 3 },
      A4: { formula: 'SUM(A1:A3)' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A4.value).toBe(6);
  });

  it('should handle IF statements in formulas', () => {
    const mockWorksheet = {
      A1: { value: 5 },
      A2: { formula: 'IF(A1>10, "Yes", "No")' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A2.value).toBe('No');
  });

  it('should ignore cells without values or formulas', () => {
    const mockWorksheet = {
      A1: {},
      B1: { value: 3 },
      C1: { formula: 'B1*2' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.C1.value).toBe(6);
  });

  // it('should handle circular references', () => {
  //   const mockWorksheet = {
  //     A1: { formula: 'A2+1' },
  //     A2: { formula: 'A1+1' },
  //   };
  //   excelCalculator.setWorksheet(mockWorksheet);
  //   const result = excelCalculator.calculate();
  //   expect(result.A1.value).toBeNull();
  //   expect(result.A2.value).toBeNull();
  // });

  it('should handle calculation 1/0 to Infinity', () => {
    const mockWorksheet = {
      A1: { formula: '1/0' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A1.value).toBe(Infinity);
  });
});
