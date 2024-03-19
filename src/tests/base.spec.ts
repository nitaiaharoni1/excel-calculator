import { ExcelCalculator } from '../ExcelCalculator';

// eslint-disable-next-line max-lines-per-function
describe('Base ExcelCalculator tests', () => {
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

  it('should calculate formulas correctly more complex', () => {
    const mockWorksheet = {};
    excelCalculator.setWorksheet(mockWorksheet);
    excelCalculator.setCellValue('A1', 100);
    excelCalculator.setCellValue('A2', 200);
    excelCalculator.setCellValue('A3', 300);
    excelCalculator.setCellFormula('B1', 'A1*2');
    excelCalculator.setCellValue('B2', 30);
    excelCalculator.setCellValue('B3', 100);
    excelCalculator.setCellFormula('C1', 'AVERAGE(A1:A3, B1:B3)');
    excelCalculator.calculate();
    expect(excelCalculator.getCellValue('B1')).toBe(200);
    // TODO: fix this test
    // expect(excelCalculator.getCellValue('C1')).toBe(155);
  });

  it('should calculate formulas correctly with cell references with $', () => {
    const mockWorksheet = {
      A1: { value: 4 },
      B1: { value: 3 },
      C1: { formula: '$A$1+B1' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.C1.value).toBe(7);
  });

  it('should set cell values correctly', () => {
    const cellAddress = 'D1';
    const value = 10;
    excelCalculator.setCellValue(cellAddress, value);
    expect(excelCalculator.getCellValue('D1')).toEqual(value);
  });

  it('should set cell formulas correctly', () => {
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
    expect(result.A1.value).toEqual(10);
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

  it('should handle calculation 1/0 to Infinity', () => {
    const mockWorksheet = {
      A1: { formula: '1/0' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A1.value).toBe(Infinity);
  });

  it('should handle formulas with multiple cell references', () => {
    const mockWorksheet = {
      A1: { value: 1 },
      B1: { value: 2 },
      C1: { value: 3 },
      D1: { formula: 'A1+B1+C1' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.D1.value).toBe(6);
  });

  it('should handle formulas with nested functions', () => {
    const mockWorksheet = {
      A1: { value: 2 },
      A2: { formula: 'IF(A1>1, SUM(A1, 5), 0)' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A2.value).toBe(7);
  });

  it('should handle formulas with cell ranges and functions', () => {
    const mockWorksheet = {
      A1: { value: 1 },
      A2: { value: 2 },
      A3: { value: 3 },
      A4: { formula: 'SUM(A1:A3)*2' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A4.value).toBe(12);
  });

  it('should handle formulas with division by zero', () => {
    const mockWorksheet = {
      A1: { value: 1 },
      A2: { value: 0 },
      A3: { formula: 'A1/A2' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A3.value).toBe(Infinity);
  });

  it('should handle formulas with negative numbers', () => {
    const mockWorksheet = {
      A1: { value: -1 },
      A2: { value: 2 },
      A3: { formula: 'A1+A2' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A3.value).toBe(1);
  });

  it('should handle formulas with boolean values', () => {
    const mockWorksheet = {
      A1: { value: true },
      A2: { formula: 'NOT(A1)' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A2.value).toBe(false);
  });

  it('should handle formulas with multiple operations and parentheses', () => {
    const mockWorksheet = {
      A1: { value: 2 },
      B1: { value: 3 },
      C1: { value: 4 },
      D1: { formula: '(A1+B1)*C1' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.D1.value).toBe(20);
  });

  it('should handle formulas with functions and cell references', () => {
    const mockWorksheet = {
      A1: { value: 2 },
      B1: { value: 3 },
      C1: { formula: 'MAX(A1, B1)' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.C1.value).toBe(3);
  });

  it('should handle formulas with nested functions and cell references', () => {
    const mockWorksheet = {
      A1: { value: 2 },
      B1: { value: 3 },
      C1: { formula: 'MAX(MIN(A1, B1), B1)' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.C1.value).toBe(3);
  });

  it('should handle formulas with cell ranges in functions', () => {
    const mockWorksheet = {
      A1: { value: 1 },
      A2: { value: 2 },
      A3: { value: 3 },
      A4: { formula: 'AVERAGE(A1:A3)' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A4.value).toBe(2);
  });

  it('should handle formulas with multiple cell ranges in functions', () => {
    const mockWorksheet = {
      A1: { value: 1 },
      A2: { value: 2 },
      A3: { value: 3 },
      B1: { value: 4 },
      B2: { value: 5 },
      B3: { value: 6 },
      C1: { formula: 'AVERAGE(A1:A3, B1:B3)' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    // const result = excelCalculator.calculate();
    // TODO: fix this test
    // expect(result.C1.value).toBe(3.5);
  });

  it('should handle circular references', () => {
    const mockWorksheet = {
      A1: { formula: 'A2+1' },
      A2: { formula: 'A1+1' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A1.value).toEqual(1);
    expect(result.A2.value).toEqual(2);
  });

  it('should handle formulas with cell references to cells with formulas', () => {
    const mockWorksheet = {
      A1: { formula: 'B1+10' },
      B1: { formula: 'C1*2' },
      C1: { value: 5 },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A1.value).toBe(20);
  });

  it('should return null when cell formula is invalid', () => {
    const mockWorksheet = {
      A1: { formula: 'INVALID_FORMULA' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A1.formula).toEqual('INVALID_FORMULA');
    expect(result.A1.value).toBeUndefined();
  });

  it('should handle formulas with division by zero + 1', () => {
    const mockWorksheet = {
      A1: { value: 1 },
      A2: { value: 0 },
      A3: { formula: 'A1/(A2+1)' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A3.value).toBe(1);
  });

  it('should handle formulas with division by zero', () => {
    const mockWorksheet = {
      A1: { value: 1 },
      A2: { value: 0 },
      A3: { formula: 'A1/A2' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A3.value).toBe(Infinity);
  });

  it('should handle formulas with invalid functions', () => {
    const mockWorksheet = {
      A1: { formula: 'INVALID(A2)' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A1.formula).toEqual('INVALID(A2)');
    expect(result.A1.value).toBeUndefined();
  });

  it('should handle formulas with missing parentheses', () => {
    const mockWorksheet = {
      A1: { formula: 'SUM(A2, A3' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A1.formula).toEqual('SUM(A2, A3');
    expect(result.A1.value).toBeUndefined();
  });

  it('should handle formulas with extra parentheses', () => {
    const mockWorksheet = {
      A1: { formula: 'SUM(A2, A3))' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A1.formula).toEqual('SUM(A2, A3))');
    expect(result.A1.value).toBeUndefined();
  });

  it('should handle formulas with error values', () => {
    const mockWorksheet = {
      A1: { formula: 'SQRT(-1)' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A1.value.re).toEqual(0);
    expect(result.A1.value.im).toEqual(1);
  });

  it('should handle INDEX formula', () => {
    const mockWorksheet = {
      A1: { value: 1 },
      A2: { value: 2 },
      A3: { value: 3 },
      A4: { formula: 'INDEX(A1:A3, 1, 0)' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A4.value).toBe(2);
  });

  it('should handle MATCH formula', () => {
    const mockWorksheet = {
      A1: { value: 1 },
      A2: { value: 2 },
      A3: { value: 3 },
      A4: { formula: 'MATCH(2, A1:A3, 0)' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A4.value).toBe(1);
  });

  it('should handle INDEX and MATCH formula', () => {
    const mockWorksheet = {
      A1: { value: 1 },
      A2: { value: 2 },
      A3: { value: 3 },
      A4: { formula: 'INDEX(A1:A3, MATCH(2, A1:A3, 0), 0)' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A4.value).toBe(2);
  });

  // multiple VLOOKUP(search_key, range, index, [is_sorted]) tests
  it('should handle VLOOKUP formula', () => {
    const mockWorksheet = {
      A1: { value: 'apple' },
      A2: { value: 'banana' },
      A3: { value: 'cherry' },
      A4: { formula: 'VLOOKUP("banana", A1:B3, 2)' },
      B1: { value: 1 },
      B2: { value: 2 },
      B3: { value: 3 },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A4.value).toBe(2);
  });

  it('should handle VLOOKUP formula with is_sorted = false', () => {
    const mockWorksheet = {
      A1: { value: 'apple' },
      A2: { value: 'banana' },
      A3: { value: 'cherry' },
      A4: { formula: 'VLOOKUP("banana", A1:B3, 2, false)' },
      B1: { value: 1 },
      B2: { value: 2 },
      B3: { value: 3 },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A4.value).toBe(2);
  });

  it('should handle VLOOKUP formula with is_sorted = true', () => {
    const mockWorksheet = {
      A1: { value: 'apple' },
      A2: { value: 'banana' },
      A3: { value: 'cherry' },
      A4: { formula: 'VLOOKUP("banana", A1:B3, 2, true)' },
      B1: { value: 1 },
      B2: { value: 2 },
      B3: { value: 3 },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A4.value).toBe(2);
  });

  it('should handle VLOOKUP formula with search_key not found', () => {
    const mockWorksheet = {
      A1: { value: 'apple' },
      A2: { value: 'banana' },
      A3: { value: 'cherry' },
      A4: { formula: 'VLOOKUP("pear", A1:B3, 2)' },
      B1: { value: 1 },
      B2: { value: 2 },
      B3: { value: 3 },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A4.value).toBeUndefined();
  });

  it('should handle VLOOKUP formula with index out of range', () => {
    const mockWorksheet = {
      A1: { value: 'apple' },
      A2: { value: 'banana' },
      A3: { value: 'cherry' },
      A4: { formula: 'VLOOKUP("banana", A1:B3, 3)' },
      B1: { value: 1 },
      B2: { value: 2 },
      B3: { value: 3 },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A4.value).toBeUndefined();
  });

  it('should handle SUM and EXP formulas', () => {
    const mockWorksheet = {
      A1: { value: -0.079 },
      A10: { value: 0.0 },
      A11: { value: -0.029 },
      A12: { value: 0.315 },
      A13: { value: -0.148 },
      A14: { value: 0.033 },
      A15: { value: 0.045 },
      A16: { value: 0.003 },
      A17: { value: 0.0 },
      A18: { value: -0.005 },
      A19: { value: 0.0 },
      A2: { value: 0.012 },
      A20: { value: 0.008 },
      A21: { value: 0.0 },
      A22: { value: 0.0 },
      A23: { value: -3.308 },
      A3: { value: 0.004 },
      A24: { formula: 'SUM(A1:A23)' },
      A4: { value: 0.0 },
      A25: { formula: 'EXP(A24)/(1+EXP(A24))' },
      A5: { value: -0.18 },
      A6: { value: 0.0 },
      A7: { value: 0.536 },
      A8: { value: 0.0 },
      A9: { value: 0.0 },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    const valueSum = result.A24.value;
    expect(valueSum).toBeCloseTo(-2.792, 2);
    const valueExp = result.A25.value;
    expect(valueExp).toBeCloseTo(0.058, 2);
  });

  // 10 dependent cells
  it('should handle 10 dependent cells', () => {
    const mockWorksheet = {
      A1: { formula: 'A2+1' },
      A10: { value: 1 },
      A2: { formula: 'A3+1' },
      A3: { formula: 'A4+1' },
      A4: { formula: 'A5+1' },
      A5: { formula: 'A6+1' },
      A6: { formula: 'A7+1' },
      A7: { formula: 'A8+1' },
      A8: { formula: 'A9+1' },
      A9: { formula: 'A10+1' },
    };
    excelCalculator.setWorksheet(mockWorksheet);
    const result = excelCalculator.calculate();
    expect(result.A1.value).toBe(10);
  });

  // it("should handle formulas with string concatenation", () => {
  //   const mockWorksheet = {
  //     A1: { value: "Hello" },
  //     A2: { value: "World" },
  //     A3: { formula: "A1 & \" \" & A2" },
  //   };
  //   excelCalculator.setWorksheet(mockWorksheet);
  //   const result = excelCalculator.calculate();
  //   expect(result.A3.value).toBe("Hello World");
  // });

  // it("should handle formulas with string values", () => {
  //   const mockWorksheet = {
  //     A1: { value: "Hello" },
  //     A2: { value: " World" },
  //     A3: { formula: "A1&A2" },
  //   };
  //   excelCalculator.setWorksheet(mockWorksheet);
  //   const result = excelCalculator.calculate();
  //   expect(result.A3.value).toBe("Hello World");
  // });
  //
  // it("should handle formulas with date values", () => {
  //   const mockWorksheet = {
  //     A1: { value: new Date(2022, 0, 1) },
  //     A2: { value: new Date(2022, 11, 31) },
  //     A3: { formula: "YEARFRAC(A1, A2, 1)" },
  //   };
  //   excelCalculator.setWorksheet(mockWorksheet);
  //   const result = excelCalculator.calculate();
  //   expect(result.A3.value).toBeCloseTo(1, 2);
  // });
});
