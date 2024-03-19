import { ExcelCalculator } from '../ExcelCalculator';
// import path
import * as path from 'path';

describe('ExcelCalculator Tests', () => {
  let excelCalculator: ExcelCalculator;

  beforeEach(async () => {
    const filePath = path.join(__dirname, '../assets/prevent.xlsx');
    const sheetName = 'Sheet1';
    excelCalculator = new ExcelCalculator(filePath, sheetName);
    await excelCalculator.init();
  });

  it('should calculate the worksheet', () => {
    excelCalculator.setCellsValues({
      D4: 1,
      D5: 54.0,
      D6: 200.0,
      D7: 50.0,
      D8: 120.0,
      D9: 0.0,
      D10: 1.0,
      D11: 25.0,
      D12: 100.0,
      D13: 1.0,
      D14: 1.0,
      D15: 200.0,
      D16: 0.87,
      D17: 2.0,
    });

    const result = excelCalculator.calculate();
    // B column
    expect(result).toBeDefined();
    expect(result.B22.value).toBeCloseTo(-0.1, 5);
    expect(result.B23.value).toBeCloseTo(0.379, 5);
    expect(result.B24.value).toBeCloseTo(-0.023333, 5);
    expect(result.B25.value).toBe(0);
    expect(result.B26.value).toBeCloseTo(-0.5, 5);
    expect(result.B27.value).toBe(0);
    expect(result.B28.value).toBe(1);
    expect(result.B29.value).toBe(0);
    expect(result.B30.value).toBe(0);
    expect(result.B31.value).toBe(0);
    expect(result.B32.value).toBeCloseTo(-0.666667, 5);
    expect(result.B33.value).toBe(1);
    expect(result.B34.value).toBe(1);
    expect(result.B35.value).toBeCloseTo(-0.5, 5);
    expect(result.B36.value).toBeCloseTo(0.379, 5);
    expect(result.B37.value).toBeCloseTo(-0.0379, 5);
    expect(result.B38.value).toBeCloseTo(0.002333, 5);
    expect(result.B39.value).toBeCloseTo(0.05, 5);
    expect(result.B40.value).toBe(0);
    expect(result.B41.value).toBeCloseTo(-0.1, 5);
    expect(result.B42.value).toBe(0);
    expect(result.B43.value).toBe(0);
    expect(result.B44.value).toBe(1);

    // C column
    expect(result.C22.value).toBeCloseTo(-0.079, 3);
    expect(result.C23.value).toBeCloseTo(0.012, 3);
    expect(result.C24.value).toBeCloseTo(0.004, 3);
    expect(result.C25.value).toBe(0);
    expect(result.C26.value).toBeCloseTo(-0.18, 3);
    expect(result.C27.value).toBe(0);
    expect(result.C28.value).toBeCloseTo(0.536, 3);
    expect(result.C29.value).toBe(0);
    expect(result.C30.value).toBe(0);
    expect(result.C31.value).toBe(0);
    expect(result.C32.value).toBeCloseTo(-0.029, 3);
    expect(result.C33.value).toBeCloseTo(0.315, 3);
    expect(result.C34.value).toBeCloseTo(-0.148, 3);
    expect(result.C35.value).toBeCloseTo(0.033, 3);
    expect(result.C36.value).toBeCloseTo(0.045, 3);
    expect(result.C37.value).toBeCloseTo(0.003, 3);
    expect(result.C38.value).toBeCloseTo(0.0, 3);
    expect(result.C39.value).toBeCloseTo(-0.005, 3);
    expect(result.C40.value).toBe(0);
    expect(result.C41.value).toBeCloseTo(0.008, 3);
    expect(result.C42.value).toBe(0);
    expect(result.C43.value).toBe(0);
    expect(result.C44.value).toBeCloseTo(-3.308, 3);
    expect(result.C45.value).toBeCloseTo(-2.792, 3);
    expect(result.C46.value).toBeCloseTo(0.058, 3);

    // D column
    expect(result.D22.value).toBeCloseTo(-0.076885, 3);
    expect(result.D23.value).toBeCloseTo(0.027901, 3);
    expect(result.D24.value).toBeCloseTo(0.002227, 3);
    expect(result.D25.value).toBe(0);
    expect(result.D26.value).toBeCloseTo(-0.168133, 3);
    expect(result.D27.value).toBe(0);
    expect(result.D28.value).toBeCloseTo(0.438687, 3);
    expect(result.D29.value).toBe(0);
    expect(result.D30.value).toBe(0);
    expect(result.D31.value).toBe(0);
    expect(result.D32.value).toBeCloseTo(-0.010988, 3);
    expect(result.D33.value).toBeCloseTo(0.288879, 3);
    expect(result.D34.value).toBeCloseTo(-0.133735, 3);
    expect(result.D35.value).toBeCloseTo(0.023796, 3);
    expect(result.D36.value).toBeCloseTo(0.056953, 3);
    expect(result.D37.value).toBeCloseTo(0.001963, 3);
    expect(result.D38.value).toBeCloseTo(0.000045, 3);
    expect(result.D39.value).toBeCloseTo(-0.005247, 3);
    expect(result.D40.value).toBe(0);
    expect(result.D41.value).toBeCloseTo(0.008951, 3);
    expect(result.D42.value).toBe(0);
    expect(result.D43.value).toBe(0);
    expect(result.D44.value).toBeCloseTo(-3.031168, 3);
    expect(result.D45.value).toBeCloseTo(-2.576755, 3);
    expect(result.D46.value).toBeCloseTo(0.070649, 3);

    // F6-J6
    expect(result.F6.value).toBeCloseTo(0.070649, 3);
    expect(result.G6.value).toBeCloseTo(0.042981, 3);
    expect(result.H6.value).toBeCloseTo(0.032242, 3);
    expect(result.I6.value).toBeCloseTo(0.023465, 3);
    expect(result.J6.value).toBeCloseTo(0.019835, 3);

    // F12-M12
    //   0.070649	Consider statins	1.150000	1.100000	1.150000	0.102777	Start statins, further evalution for CVD	85.000000
    expect(result.F12.value).toBeCloseTo(0.070649, 3);
    expect(result.H12.value).toBeCloseTo(1.15, 3);
    expect(result.I12.value).toBeCloseTo(1.1, 3);
    expect(result.J12.value).toBeCloseTo(1.15, 3);
    expect(result.K12.value).toBeCloseTo(0.102777, 3);
    expect(result.M12.value).toBeCloseTo(85.0, 3);
  });
});
