// spreadsheet.service.ts
import { Injectable } from '@angular/core';
import { SheetData, SpreadsheetFile, FileFormat, CellData, CellStyle } from './spreadsheet.types';
import * as XLSX from 'xlsx';

@Injectable({
  providedIn: 'root'
})
export class SpreadsheetService {
  async saveFile(data: SpreadsheetFile, format: FileFormat): Promise<boolean> {
    try {
      const workbook = XLSX.utils.book_new();

      data.sheets.forEach(sheet => {
        // Convert CellData[][] to format suitable for XLSX
        const wsData = sheet.data.map(row =>
          row.map(cell => ({
            v: cell.value,
            s: this.convertStyleToXLSX(cell.style)
          }))
        );

        const ws = XLSX.utils.aoa_to_sheet(wsData);
        XLSX.utils.book_append_sheet(workbook, ws, sheet.name);
      });

      // Use browser's file save dialog
      const fileName = `spreadsheet.${format}`;
      XLSX.writeFile(workbook, fileName);
      return true;
    } catch (error) {
      console.error('Error saving file:', error);
      return false;
    }
  }

  async openFile(format: FileFormat): Promise<SpreadsheetFile | null> {
    try {
      // Create file input element
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = `.${format}`;

      return new Promise((resolve) => {
        input.onchange = async (e: Event) => {
          const target = e.target as HTMLInputElement;
          const file = target.files?.[0];

          if (!file) {
            resolve(null);
            return;
          }

          const reader = new FileReader();
          reader.onload = (e) => {
            const data = e.target?.result;
            const workbook = XLSX.read(data, { type: 'binary' });

            const sheets: SheetData[] = workbook.SheetNames.map((name, index) => {
              const sheet = workbook.Sheets[name];
              const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

              // Convert to CellData format
              const cellData: CellData[][] = (jsonData as any[][]).map((row: any[]) =>
                row.map((cell: any, colIndex: number) => ({
                  value: cell,
                  style: this.extractStyleFromXLSX(
                    sheet[XLSX.utils.encode_cell({ r: jsonData.indexOf(row), c: colIndex })]
                  )
                }))
              );

              return {
                name,
                data: cellData,
                id: index,
                columnHeaders: this.generateColumnHeaders(cellData[0]?.length || 26)
              };
            });

            resolve({
              sheets,
              activeSheet: 0
            });
          };

          reader.readAsBinaryString(file);
        };

        input.click();
      });
    } catch (error) {
      console.error('Error opening file:', error);
      return null;
    }
  }

  private generateColumnHeaders(length: number): string[] {
    return Array.from({ length }, (_, i) => String.fromCharCode(65 + i));
  }

  private convertStyleToXLSX(style?: CellStyle): any {
    if (!style) return {};

    return {
      font: {
        bold: style.bold,
        italic: style.italic,
        underline: style.underline
      },
      fill: style.backgroundColor ? {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { rgb: style.backgroundColor.replace('#', '') }
      } : undefined
    };
  }

  private extractStyleFromXLSX(cell: any): CellStyle {
    if (!cell || !cell.s) return {};

    return {
      bold: cell.s.font?.bold,
      italic: cell.s.font?.italic,
      underline: cell.s.font?.underline,
      backgroundColor: cell.s.fill?.fgColor?.rgb ? `#${cell.s.fill.fgColor.rgb}` : undefined
    };
  }
}
