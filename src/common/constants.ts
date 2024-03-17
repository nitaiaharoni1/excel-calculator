export const MULTIPLE_ARGS_FORMULAS = ['INDEX', 'MATCH', 'VLOOKUP'];
export const EXCEL_TO_MATHJS_FORMULAS = {
  ABS: 'abs',
  ACOS: 'acos',
  ACOSH: 'acosh',
  AND: 'and', // Note: In math.js, 'and' is bitwise; might need custom logic for array inputs
  ASIN: 'asin',
  ASINH: 'asinh',
  ATAN: 'atan',
  ATAN2: 'atan2', // Added: Excel's ATAN2 function
  ATANH: 'atanh',
  AVERAGE: 'mean',
  CEILING: 'ceil',
  COS: 'cos',
  COSH: 'cosh',
  COUNT: 'size', // Corrected: 'length' in JavaScript gets an array's length, 'size' in math.js can count matrix/array elements
  EXP: 'exp',
  FLOOR: 'floor',
  INT: 'floor', // Added: Excel's INT function, similar to FLOOR for positive numbers
  LOG: 'log',
  LOG10: 'log10',
  MAX: 'max',
  MIN: 'min',
  MOD: 'mod', // Added: Excel's MOD function
  NOT: 'not', // Note: In math.js, 'not' is bitwise; might need custom logic for array inputs
  OR: 'or', // Note: In math.js, 'or' is bitwise; might need custom logic for array inputs
  PI: () => 'pi', // Added: Excel's PI function, Note: Use as a function in your replacement logic
  POWER: 'pow',
  RADIANS: 'to', // Added: Convert degrees to radians using 'to' function with 'radian' as the unit in math.js
  RAND: 'random', // Added: Excel's RAND function
  ROUND: 'round',
  ROUNDDOWN: 'floor', // Added: Excel's ROUNDDOWN function
  ROUNDUP: 'ceil', // Added: Excel's ROUNDUP function
  SIN: 'sin',
  SINH: 'sinh',
  SQRT: 'sqrt',
  SUM: 'sum',
  TAN: 'tan',
  TANH: 'tanh',
  TRUNC: 'fix', // Added: Excel's TRUNC function maps to 'fix' in math.js
};
