export class FormulaParser {
  static parseCellRange(range: string): { start: { row: number; col: string }, end: { row: number; col: string } } {
    const [start, end] = range.split(':').map(cell => this.parseCellReference(cell));
    return { start, end };
  }

  static parseCellReference(ref: string): { row: number; col: string } {
    const match = ref.match(/([A-Z]+)(\d+)/);
    if (!match) throw new Error('Invalid cell reference');
    const col = match[1];
    const row = parseInt(match[2]) - 1;
    return { row, col };
  }

  static getRangeValues(data: any[][], columnHeaders: string[], range: string): number[] {
    const { start, end } = this.parseCellRange(range);
    const startColIndex = columnHeaders.indexOf(start.col);
    const endColIndex = columnHeaders.indexOf(end.col);

    const values: number[] = [];

    for (let row = start.row; row <= end.row; row++) {
      for (let col = startColIndex; col <= endColIndex; col++) {
        const value = this.extractNumericValue(data[row][col]);
        if (!isNaN(value)) {
          values.push(value);
        }
      }
    }

    return values;
  }

  static extractNumericValue(cellContent: any): number {
    if (typeof cellContent === 'number') return cellContent;
    if (typeof cellContent === 'string') {
      if (cellContent.startsWith('=')) {
        return NaN;
      }
      const parsed = parseFloat(cellContent);
      return isNaN(parsed) ? NaN : parsed;
    }
    return NaN;
  }

  // Mathematical Functions
  static SUM(data: any[][], columnHeaders: string[], range: string): number {
    const values = this.getRangeValues(data, columnHeaders, range);
    return values.reduce((sum, val) => sum + val, 0);
  }

  static AVG(data: any[][], columnHeaders: string[], range: string): number {
    const values = this.getRangeValues(data, columnHeaders, range);
    return values.length > 0 ? values.reduce((sum, val) => sum + val, 0) / values.length : 0;
  }

  static MAX(data: any[][], columnHeaders: string[], range: string): number {
    const values = this.getRangeValues(data, columnHeaders, range);
    return values.length > 0 ? Math.max(...values) : 0;
  }

  static MIN(data: any[][], columnHeaders: string[], range: string): number {
    const values = this.getRangeValues(data, columnHeaders, range);
    return values.length > 0 ? Math.min(...values) : 0;
  }

  static COUNT(data: any[][], columnHeaders: string[], range: string): number {
    const values = this.getRangeValues(data, columnHeaders, range);
    return values.length;
  }

  static TRIM(value: string): string {
    return value.trim();
  }

  static UPPER(value: string): string {
    return value.toUpperCase();
  }

  static LOWER(value: string): string {
    return value.toLowerCase();
  }

  static REMOVE_DUPLICATES(data: any[][], columnHeaders: string[], range: string): any[][] {
    const { start, end } = this.parseCellRange(range);
    const startColIndex = columnHeaders.indexOf(start.col);
    const endColIndex = columnHeaders.indexOf(end.col);

    const rangeMap = new Map<string, number>();
    const uniqueRows: number[] = [];

    for (let row = start.row; row <= end.row; row++) {
      const rowValues = data[row].slice(startColIndex, endColIndex + 1);
      const rowKey = JSON.stringify(rowValues);

      if (!rangeMap.has(rowKey)) {
        rangeMap.set(rowKey, row);
        uniqueRows.push(row);
      }
    }

    // Return only the unique rows within the specified range
    return uniqueRows.map(row => {
      const newRow = [...data[row]];
      // Only modify the cells within the specified range
      for (let col = startColIndex; col <= endColIndex; col++) {
        newRow[col] = data[row][col];
      }
      return newRow;
    });
  }

  static FIND_AND_REPLACE(data: any[][], columnHeaders: string[], range: string, find: string, replace: string): any[][] {
    const { start, end } = this.parseCellRange(range);
    const startColIndex = columnHeaders.indexOf(start.col);
    const endColIndex = columnHeaders.indexOf(end.col);

    // Create a deep copy of the data
    const newData = data.map(row => [...row]);

    // Remove quotes if they exist
    const findStr = find.replace(/^"(.*)"$/, '$1');
    const replaceStr = replace.replace(/^"(.*)"$/, '$1');

    for (let row = start.row; row <= end.row; row++) {
      for (let col = startColIndex; col <= endColIndex; col++) {
        if (typeof newData[row][col] === 'string') {
          newData[row][col] = newData[row][col].replace(new RegExp(findStr, 'g'), replaceStr);
        }
      }
    }

    return newData;
  }

  static splitArguments(args: string): string[] {
    const result: string[] = [];
    let current = '';
    let inQuotes = false;
    let depth = 0;

    for (const char of args) {
      if (char === '"' && depth === 0) {
        inQuotes = !inQuotes;
        current += char;
      } else if (char === ',' && !inQuotes && depth === 0) {
        result.push(current.trim());
        current = '';
      } else {
        if (char === '(') depth++;
        if (char === ')') depth--;
        current += char;
      }
    }

    if (current) {
      result.push(current.trim());
    }

    return result;
  }

  // Main formula parser
  static parseFormula(formula: string, data: any[][], columnHeaders: string[]): any {
    formula = formula.trim();

    if (!formula.startsWith('=')) return formula;

    const formulaPattern = /^=([A-Z_]+)\((.*)\)$/;
    const match = formula.match(formulaPattern);

    if (!match) return '#ERROR!';

    const [_, functionName, args] = match;
    const argumentList = this.splitArguments(args);

    try {
      switch (functionName) {
        case 'SUM':
          return this.SUM(data, columnHeaders, argumentList[0]);
        case 'AVG':
          return this.AVG(data, columnHeaders, argumentList[0]);
        case 'MAX':
          return this.MAX(data, columnHeaders, argumentList[0]);
        case 'MIN':
          return this.MIN(data, columnHeaders, argumentList[0]);
        case 'COUNT':
          return this.COUNT(data, columnHeaders, argumentList[0]);
        case 'TRIM':
          return this.TRIM(argumentList[0]);
        case 'UPPER':
          return this.UPPER(argumentList[0]);
        case 'LOWER':
          return this.LOWER(argumentList[0]);
        case 'REMOVE_DUPLICATES':
          return this.REMOVE_DUPLICATES(data, columnHeaders, argumentList[0]);
        case 'FIND_AND_REPLACE':
          if (argumentList.length !== 3) return '#ERROR: Requires 3 arguments!';
          return this.FIND_AND_REPLACE(data, columnHeaders, argumentList[0], argumentList[1], argumentList[2]);
        default:
          return '#INVALID!';
      }
    } catch (error) {
      console.error('Formula parsing error:', error);
      return '#ERROR!';
    }
  }

  static parseFormulaWithCellRef(input: string, data: any[][], columnHeaders: string[]): {
    targetCell: { row: number; col: string } | null;
    formula: string | null;
  } {
    // Match patterns like "A1=SUM(B1:B8)" or just "=SUM(B1:B8)"
    const cellAssignmentPattern = /^([A-Z]+\d+)=(.+)$/;
    const match = input.match(cellAssignmentPattern);

    if (match) {
      // We have a cell assignment like "A1=SUM(B1:B8)"
      const [_, targetCellRef, formulaContent] = match;
      const targetCell = this.parseCellReference(targetCellRef);
      const formula = formulaContent.startsWith('=') ? formulaContent : `=${formulaContent}`;
      return { targetCell, formula };
    } else {
      // Regular formula without cell assignment
      return { targetCell: null, formula: input };
    }
  }

  static evaluateFormula(formula: string, data: any[][], columnHeaders: string[]): any {
    try {
      // Remove the '=' if it exists at the start
      const cleanFormula = formula.startsWith('=') ? formula : `=${formula}`;
      return this.parseFormula(cleanFormula, data, columnHeaders);
    } catch (error) {
      console.error('Error evaluating formula:', error);
      return '#ERROR!';
    }
  }
}




