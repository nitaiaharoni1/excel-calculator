import * as ExcelJS from 'exceljs';
import { IWorksheet } from '../types';

export function convertIfToTernary(formula: string): string {
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
  const parts = parseIfContent(ifContent);

  if (parts.length === 3) {
    const ternaryExpression = `${parts[0]} ? ${parts[1]} : ${parts[2]}`;
    return formula.substring(0, startIndex) + ternaryExpression + formula.substring(endIndex);
  }
  return formula; // In case of unexpected format
}

function parseIfContent(ifContent: string): string[] {
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

export function buildWorksheet(excelWorksheet: ExcelJS.Worksheet): IWorksheet {
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

export function customIndexFunction(array: any[][], rowIndex: number, columnIndex: number = 1): any {
  // @ts-expect-error
  const parsedArray = array._data;
  if (Array.isArray(parsedArray) && rowIndex > 0 && columnIndex > 0) {
    return parsedArray[rowIndex - 1]?.[columnIndex - 1] ?? null;
  }
  throw new Error('INDEX function parameters out of range');
}

export function customMatchFunction(lookupValue: any, lookupArray: any[], matchType: number = 0): number | null {
  // @ts-expect-error
  const parsedLookupArray = lookupArray._data.flat();
  if (Array.isArray(parsedLookupArray)) {
    const index = parsedLookupArray.findIndex((item) => item === lookupValue);
    return index >= 0 ? index + 1 : null; // Excel is 1-based index; return null if not found
  }
  throw new Error('MATCH function lookupArray must be an array');
}

// VLOOKUP(search_key, range, index, [is_sorted])
export function customVlookupFunction(searchKey: any, range: any[][], index: number, isSorted: boolean = true): any | null {
  // @ts-expect-error
  const parsedRange = range._data;
  if (Array.isArray(parsedRange) && index > 0) {
    const searchIndex = parsedRange.findIndex((row) => row[0] === searchKey);
    if (searchIndex >= 0) {
      return parsedRange[searchIndex][index - 1];
    }
  }
  throw new Error('VLOOKUP function parameters out of range');
}
