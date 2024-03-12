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
