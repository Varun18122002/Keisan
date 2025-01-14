
import { Component, HostListener, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HyperFormula } from 'hyperformula';
import { RightclickComponent } from '../rightclick/rightclick.component';
import { ToolbarComponent } from "../toolbar/toolbar.component";
import { FormulaParser } from '../formula-functions/formula-functions';
import { DataValidator } from '../formula-functions/datavalid';

interface SortState {
  column: string;
  direction: 'asc' | 'desc' | null;
  lastClickTime: number;
}

export interface CellMetadata {
  type: 'text' | 'number' | 'date';
  validation?: {
    isValid: boolean;
    message?: string;
  };
}


@Component({
  selector: 'app-grid',
  standalone: true,
  imports: [CommonModule, RightclickComponent, ToolbarComponent],
  template: `
<app-toolbar
  [gridData]="displayData"
  [selectedRange]="selectedRange"
  [currentFormula]="currentFormula"
  (formulaInput)="onFormulaBarInput($event)"
  (formulaEnter)="onFormulaEnter($event)"
  (formulaFocus)="onFormulaFocus()"
  (formulaBlur)="onFormulaBlur()"
  (fileAction)="handleFileAction($event)"
    />
    <div class="sheet-container"
         (mouseup)="onMouseUp()"
         (mouseleave)="onMouseLeave()">

      <div class="sheet-tabs">

          <button
            *ngFor="let sheet of sheets; let i = index"
            [class.active]="i === activeSheetIndex"
            (click)="selectSheet(i)"
          >
            {{ sheet.name }}
            <span class="remove-sheet-btn" (click)="removeSheet(i, $event)">&#45;</span>
          </button>
          <button class="add-sheet-btn" (click)="addSheet()">+</button>

      </div>

      <div class="grid-container">
        <table class="spreadsheet">
        <thead>
          <tr>
            <th>#</th>
            <ng-container *ngFor="let col of sheets[activeSheetIndex].columnHeaders; let colIndex = index">
            <th class="column-header" [style.width.px]="columnWidths[colIndex] || 100"
    draggable="true"
    (dragstart)="onDragStart($event, col, 'column')"
    (dragover)="onDragOver($event)"
    (drop)="onDrop($event, col, 'column')"
    (dragend)="onDragEnd()"
    [class.dragging]="isDraggingColumn && draggedColumn === col">
    <div class="header-content"
         (click)="handleColumnHeaderClick($event, col)">  <!-- Add this click handler -->
        {{ col }}
        <span class="sort-indicator" *ngIf="sortState.column === col">
            {{ sortState.direction === 'asc' ? '▲' : '▼' }}
        </span>
    </div>
    <div class="column-resize-handle"
         (mousedown)="onColumnResizerMouseDown($event, colIndex)">
    </div>
</th>
            </ng-container>
          </tr>
        </thead>
      <tbody>
        <tr *ngFor="let row of displayData; let rowIndex = index"
        [style.height.px]="rowHeights[rowIndex] || 25">
        <td class="row-header"
                  draggable="true"
                  (dragstart)="onDragStart($event, rowIndex.toString(), 'row')"
                  (dragover)="onDragOver($event)"
                  (drop)="onDrop($event, rowIndex.toString(), 'row')"
                  (dragend)="onDragEnd()"
                  [class.dragging]="isDraggingRow && draggedRow === rowIndex.toString()">
                {{ rowIndex + 1 }}
                <div
                  class="row-resize-handle"
                  (mousedown)="onRowResizerMouseDown($event, rowIndex)"
                ></div>
              </td>
              <!-- Data cells -->

      <ng-container *ngFor="let col of sheets[activeSheetIndex].columnHeaders">
        <td
                  [attr.data-row]="rowIndex"
                  [attr.data-col]="col"
                  [ngStyle]="getCellStyle(rowIndex, col)"
                  (mousedown)="onCellMouseDown($event, rowIndex, col)"
                  (mouseover)="onCellMouseOver($event, rowIndex, col)"
                  (click)="onCellClick($event, rowIndex, col)"
                  [class.selected]="isCellSelected(rowIndex, col)"
                  contenteditable="true"
                  (input)="onCellEdit($event, rowIndex, col)"
                  (blur)="onCellBlur($event, rowIndex, col)"
                  (focus)="onCellFocus($event, rowIndex, col)"
        >
          {{ getCellValue(rowIndex, col) }}
        </td>
      </ng-container>
    </tr>
  </tbody>
</table>
        <app-rightclick
          [isVisible]="showContextMenu"
          [position]="contextMenuPosition"
          [selectedCells]="selectedRange"
          (menuAction)="handleMenuAction($event)"
        ></app-rightclick>
      </div>
    </div>
  `,

  styleUrl: 'grid.component.scss',
})

export class GridComponent implements OnInit {

  private hf: HyperFormula;

  private isDirectCellEdit = false;

  cellMetadata: Map<string, CellMetadata> = new Map();

  activeSheetIndex = 0;
  sheets: Array<{
    name: string;
    data: any[][];
    id: number;
    columnHeaders: string[];
  }> = [];

  private clipboardData: {
    data: any[][];
    startCol: string;
    endCol: string;
    startRow: number;
    endRow: number;
  } | null = null;


  displayData: any[][] = [];
  showContextMenu = false;
  contextMenuPosition = { x: 0, y: 0 };
  selectedCells: any = null;
  selectedCell: { rowIndex: number, colId: string } | null = null;
  draggedColumn: string | null = null;
  isDraggingColumn = false;
  draggedRow: string | null = null;
  isDraggingRow = false;
  isSelecting = false;
  selectionStart: { row: number; col: string } | null = null;
  selectionEnd: { row: number; col: string } | null = null;
  selectedRange: {
    startRow: number;
    endRow: number;
    startCol: string;
    endCol: string;
  } | null = null;

  isDraggingCells = false;
  draggedCellsData: any[][] = [];
  dragStartCell: { row: number; col: string } | null = null;

  private readonly STORAGE_KEY = 'spreadsheet_data';

  cellFormatting: Map<string, {
    bold: boolean;
    italic: boolean;
    fontSize: number;
    color: string;
  }> = new Map();

  isResizingColumn = false;
  isResizingRow = false;
  resizeStartPosition = 0;
  resizeColumnIndex: number | null = null;
  resizeRowIndex: number | null = null;
  columnWidths: number[] = [];
  rowHeights: number[] = [];

  onDragEnd(): void {
    this.isDraggingColumn = false;
    this.draggedColumn = null;
    this.isDraggingRow = false;
    this.draggedRow = null;
    this.isDraggingCells = false;
    this.dragStartCell = null;
    this.updateDisplayData();
  }

  getCellFormatting(rowIndex: number, colId: string): any {
    const key = `${rowIndex}-${colId}`;
    return this.cellFormatting.get(key) || {
      bold: false,
      italic: false,
      fontSize: 12,
      color: 'black'
    };
  }


  onFormulaFocus() {
    console.log('Formula bar focused');
    this.isFormulaBarFocused = true;
  }

  onFormulaBlur() {
    console.log('Formula bar blur');
    this.isFormulaBarFocused = false;
  }

  // onFormulaBarInput(event: Event) {
  //   const input = event.target as HTMLInputElement;
  //   this.currentFormula = input.value;
  //   console.log('Formula bar input:', this.currentFormula);
  // }


  updateCellFormatting(rowIndex: number, colId: string, formatting: any) {
    const key = `${rowIndex}-${colId}`;
    this.cellFormatting.set(key, {
      ...this.getCellFormatting(rowIndex, colId),
      ...formatting
    });
    this.updateDisplayData();
  }

  onColumnResizerMouseDown(event: MouseEvent, columnIndex: number) {
    event.preventDefault();
    this.isResizingColumn = true;
    this.resizeColumnIndex = columnIndex;
    this.resizeStartPosition = event.clientX;
  }

  onRowResizerMouseDown(event: MouseEvent, rowIndex: number) {
    event.preventDefault();
    this.isResizingRow = true;
    this.resizeRowIndex = rowIndex;
    this.resizeStartPosition = event.clientY;
  }

  onMouseUp() {
    this.isSelecting = false;
    this.isResizingColumn = false;
    this.isResizingRow = false;
  }

  onFormatChange(event: { type: string, value: boolean }) {
    if (!this.selectedRange) return;

    const { startRow, endRow, startCol, endCol } = this.selectedRange;
    const currentSheet = this.sheets[this.activeSheetIndex];
    const startColIndex = currentSheet.columnHeaders.indexOf(startCol);
    const endColIndex = currentSheet.columnHeaders.indexOf(endCol);
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startColIndex; col <= endColIndex; col++) {
        const colId = currentSheet.columnHeaders[col];
        this.updateCellFormatting(row, colId, {
          [event.type]: event.value
        });
      }
    }
  }

  private clearSelection() {
    this.selectedRange = null;
    this.selectionStart = null;
    this.selectionEnd = null;
    this.isDraggingCells = false;
    this.draggedCellsData = [];
  }

  constructor() {
    this.hf = HyperFormula.buildEmpty({
      licenseKey: 'gpl-v3',
    });
    this.loadFromStorage();
  }

  private loadFromStorage() {
    try {
      const savedData = localStorage.getItem(this.STORAGE_KEY);
      if (savedData) {
        const parsedData = JSON.parse(savedData);
        if (this.isValidStoredData(parsedData)) {
          this.sheets = parsedData.sheets;
          this.activeSheetIndex = parsedData.activeSheetIndex;
          this.sheets.forEach(sheet => {
            const sheetId = this.hf.addSheet(sheet.name);
            sheet.id = typeof sheetId === 'number' ? sheetId : sheet.id;
          });
        } else {
          this.initializeSheets();
        }
      } else {
        this.initializeSheets();
      }
    } catch (error) {
      console.error('Error loading from storage:', error);
      this.initializeSheets();
    }
    this.updateDisplayData();
  }

  private isValidStoredData(data: any): boolean {
    return (
      data &&
      Array.isArray(data.sheets) &&
      data.sheets.length > 0 &&
      data.sheets.every((sheet: any) =>
        sheet.name &&
        Array.isArray(sheet.data) &&
        Array.isArray(sheet.columnHeaders) &&
        typeof sheet.id === 'number'
      ) &&
      typeof data.activeSheetIndex === 'number' &&
      data.activeSheetIndex >= 0 &&
      data.activeSheetIndex < data.sheets.length
    );
  }

  private saveToStorage() {
    try {
      const dataToSave = {
        sheets: this.sheets,
        activeSheetIndex: this.activeSheetIndex
      };
      localStorage.setItem(this.STORAGE_KEY, JSON.stringify(dataToSave));
    } catch (error) {
      console.error('Error saving to storage:', error);
    }
  }

  private initializeSheets() {

    const defaultColumnHeaders = this.generateColumnHeaders();
    const sheet1Id = this.hf.addSheet('Sheet 1');
    const sheet2Id = this.hf.addSheet('Sheet 2');
    this.sheets = [
      {
        name: 'Sheet 1',
        data: this.generateInitialData(),
        id: typeof sheet1Id === 'number' ? sheet1Id : 0,
        columnHeaders: [...defaultColumnHeaders]
      },
      {
        name: 'Sheet 2',
        data: this.generateInitialData(),
        id: typeof sheet2Id === 'number' ? sheet2Id : 1,
        columnHeaders: [...defaultColumnHeaders]
      }
    ];

    this.updateDisplayData();
  }

  private generateColumnHeaders(): string[] {
    return Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i));
  }

  private numberToColumnName(num: number): string {
    let columnName = '';
    do {
      const remainder = num % 26;
      columnName = String.fromCharCode(65 + remainder) + columnName;
      num = Math.floor(num / 26) - 1;
    } while (num >= 0);
    return columnName;
  }

  private updateColumnNames() {
    const currentSheet = this.sheets[this.activeSheetIndex];
    if (!currentSheet) return;

    const totalColumns = currentSheet.columnHeaders.length;
    const newHeaders = Array.from(
      { length: totalColumns },
      (_, index) => this.numberToColumnName(index)
    );

    currentSheet.columnHeaders = newHeaders;
    currentSheet.data = currentSheet.data.map(row => {
      const newRow = [...row];
      while (newRow.length < totalColumns) {
        newRow.push('');
      }
      if (newRow.length > totalColumns) {
        newRow.length = totalColumns;
      }
      return newRow;
    });
  }

  private generateInitialData(): any[][] {
    return Array.from({ length: 100 }, () =>
      Array.from({ length: 26 }, () => '')
    );
  }

  private updateDisplayData() {

    if (this.sheets[this.activeSheetIndex]) {
      this.displayData = [...this.sheets[this.activeSheetIndex].data];
    } else {
      this.displayData = [];
    }
    this.saveToStorage();
  }

  getCellValue(rowIndex: number, colId: string): string {
    const currentSheet = this.sheets[this.activeSheetIndex];
    const colIndex = currentSheet.columnHeaders.indexOf(colId);
    const value = currentSheet.data[rowIndex][colIndex];

    if (typeof value === 'string' &&
      value.startsWith('=')) {
      try {
        const result = FormulaParser.parseFormula(
          value,
          currentSheet.data,
          currentSheet.columnHeaders
        );
        return result.toString();
      } catch (error) {
        console.error('Error calculating formula:', error);
        return '#ERROR!';
      }
    }

    return value || '';
  }



  onDragStart(event: DragEvent, item: string, type: 'row' | 'column') {
    if (event.dataTransfer) {
      if (type === 'column') {
        this.draggedColumn = item;
        this.isDraggingColumn = true;
      } else {
        this.draggedRow = item;
        this.isDraggingRow = true;
      }
      event.dataTransfer.setData('text', item);
      event.dataTransfer.setData('type', type);
    }
  }

  onDragOver(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
  }

  onDrop(event: DragEvent, target: string, type: 'row' | 'column') {
    event.preventDefault();
    event.stopPropagation();

    const dragType = event.dataTransfer?.getData('type');
    if (dragType !== type) return;

    if (type === 'column') {
      this.handleColumnDrop(target);
    } else {
      this.handleRowDrop(target);
    }
  }

  clearStorage() {
    try {
      localStorage.removeItem(this.STORAGE_KEY);
      this.initializeSheets();
      this.updateDisplayData();
    } catch (error) {
      console.error('Error clearing storage:', error);
    }
  }

  private handleRowDrop(targetRow: string) {
    if (!this.draggedRow || targetRow === this.draggedRow) {
      return;
    }

    const fromIndex = parseInt(this.draggedRow);
    const toIndex = parseInt(targetRow);
    const currentSheet = this.sheets[this.activeSheetIndex];
    const newData = [...currentSheet.data];
    const [movedRow] = newData.splice(fromIndex, 1);
    newData.splice(toIndex, 0, movedRow);
    this.sheets[this.activeSheetIndex] = {
      ...currentSheet,
      data: newData
    };

    this.updateDisplayData();
  }
  private handleColumnDrop(targetColumn: string) {
    if (!this.draggedColumn || targetColumn === this.draggedColumn) {
      return;
    }

    const currentSheet = this.sheets[this.activeSheetIndex];
    const fromIndex = currentSheet.columnHeaders.indexOf(this.draggedColumn);
    const toIndex = currentSheet.columnHeaders.indexOf(targetColumn);
    const newData = currentSheet.data.map(row => {
      const newRow = [...row];
      const [movedValue] = newRow.splice(fromIndex, 1);
      newRow.splice(toIndex, 0, movedValue);
      return newRow;
    });
    const newHeaders = [...currentSheet.columnHeaders];
    newHeaders.splice(fromIndex, 1);
    newHeaders.splice(toIndex, 0, this.draggedColumn);
    this.sheets[this.activeSheetIndex] = {
      ...currentSheet,
      columnHeaders: newHeaders,
      data: newData
    };
    this.updateColumnNames();
    this.updateDisplayData();
  }

  getCurrentSheetColumnHeaders(): string[] {
    return this.sheets[this.activeSheetIndex]?.columnHeaders || [];
  }

  addSheet() {
    try {
      const newSheetName = `Sheet ${this.sheets.length + 1}`;
      const newSheetId = this.hf.addSheet(newSheetName);

      const newSheet = {
        name: newSheetName,
        data: this.generateInitialData(),
        id: typeof newSheetId === 'number' ? newSheetId : this.sheets.length,
        columnHeaders: this.generateColumnHeaders()
      };

      this.sheets = [...this.sheets, newSheet];
      this.activeSheetIndex = this.sheets.length - 1;
      this.updateDisplayData();
    } catch (error) {
      console.error('Error adding new sheet:', error);
    }
  }
  private editingCell: HTMLElement | null = null;

  onFormulaInput(event: Event) {
    const input = event.target as HTMLInputElement;
    this.currentFormula = input.value;
  }

  onFormulaKeypress(event: KeyboardEvent) {
    if (event.key === 'Enter') {
      this.applyFormula();
    }
  }

  private applyFormula() {
    if (!this.selectedCell) return;

    const { rowIndex, colId } = this.selectedCell;
    const currentSheet = this.sheets[this.activeSheetIndex];
    const colIndex = currentSheet.columnHeaders.indexOf(colId);

    // Store the formula in the data
    currentSheet.data[rowIndex][colIndex] = this.currentFormula;

    // If it's a formula, calculate and display the result
    if (this.currentFormula.startsWith('=')) {
      const result = FormulaParser.parseFormula(
        this.currentFormula,
        currentSheet.data,
        currentSheet.columnHeaders
      );

      // Update the cell display with the calculated result
      const cell = document.querySelector(
        `td[data-row="${rowIndex}"][data-col="${colId}"]`
      ) as HTMLElement;

      if (cell) {
        cell.innerText = result.toString();
      }
    }

    this.updateDisplayData();
  }

  onCellBlur(event: FocusEvent, rowIndex: number, colId: string) {
    const target = event.target as HTMLElement;
    const value = target.textContent?.trim() || '';
    this.isDirectCellEdit = false;
    this.updateCellValue(rowIndex, colId, value);
  }

  onCellFocus(event: FocusEvent, rowIndex: number, colId: string): void {
    const currentSheet = this.sheets[this.activeSheetIndex];
    const colIndex = currentSheet.columnHeaders.indexOf(colId);
    const rawValue = currentSheet.data[rowIndex][colIndex];

    this.selectedCell = { rowIndex, colId };
    this.currentFormula = rawValue || '';

    // Update formula bar with raw value (including formula)
    if (this.formulaInput) {
      this.formulaInput.value = rawValue || '';
    }

    // Handle formula cells differently
    if (typeof rawValue === 'string' && rawValue.startsWith('=')) {
      event.preventDefault();
      (event.target as HTMLElement).blur();
      return;
    }
  }

  private handleFormulaCell(cell: HTMLElement, formula: string): void {
    // Keep the calculated value visible
    const result = this.evaluateFormula(formula);
    cell.textContent = result;

    // Make cell non-editable
    cell.contentEditable = 'false';
    // Re-enable editability after a short delay (for non-formula edits)
    setTimeout(() => {
      cell.contentEditable = 'true';
    }, 100);
  }

  private evaluateFormula(formula: string): string {
    try {
      const currentSheet = this.sheets[this.activeSheetIndex];
      const result = FormulaParser.parseFormula(
        formula,
        currentSheet.data,
        currentSheet.columnHeaders
      );
      return result.toString();
    } catch (error) {
      console.error('Formula parsing error:', error);
      return '#ERROR!';
    }
  }



  @HostListener('window:mousemove', ['$event'])
  onMouseMove(event: MouseEvent) {
    if (this.isResizingColumn && this.resizeColumnIndex !== null) {
      event.preventDefault();
      const delta = event.clientX - this.resizeStartPosition;
      const newWidth = (this.columnWidths[this.resizeColumnIndex] || 100) + delta;

      if (newWidth >= 50) {
        this.columnWidths[this.resizeColumnIndex] = newWidth;
        this.resizeStartPosition = event.clientX;
      }
    }

    if (this.isResizingRow && this.resizeRowIndex !== null) {
      event.preventDefault();
      const delta = event.clientY - this.resizeStartPosition;
      const newHeight = (this.rowHeights[this.resizeRowIndex] || 25) + delta;

      if (newHeight >= 20) {
        this.rowHeights[this.resizeRowIndex] = newHeight;
        this.resizeStartPosition = event.clientY;
      }
    }
  }
  getCellStyle(rowIndex: number, colId: string) {
    const formatting = this.getCellFormatting(rowIndex, colId);
    const columnIndex = this.getCurrentSheetColumnHeaders().indexOf(colId);

    return {
      'font-weight': formatting.bold ? 'bold' : 'normal',
      'font-style': formatting.italic ? 'italic' : 'normal',
      'font-size': `${formatting.fontSize || 14}px`,
      'color': formatting.color || '#333',
      'width': `${this.columnWidths[columnIndex] || 100}px`,
      'height': `${this.rowHeights[rowIndex] || 25}px`,
      'text-align': 'right',
      'padding': '8px',
      'white-space': 'nowrap',
      'overflow': 'hidden',
      'text-overflow': 'ellipsis'
    };
  }




  removeSheet(index: number, event: MouseEvent) {
    event.stopPropagation();

    if (this.sheets.length > 1) {
      try {
        const sheetToRemove = this.sheets[index];
        this.hf.removeSheet(sheetToRemove.id);

        this.sheets = this.sheets.filter((_, i) => i !== index);
        this.sheets = this.sheets.map((sheet, i) => ({
          ...sheet,
          name: `Sheet ${i + 1}`
        }));

        if (this.activeSheetIndex >= this.sheets.length) {
          this.activeSheetIndex = this.sheets.length - 1;
        } else if (index === this.activeSheetIndex && this.activeSheetIndex > 0) {
          this.activeSheetIndex--;
        }
        this.updateDisplayData();
      } catch (error) {
        console.error('Error removing sheet:', error);
      }
    }
  }

  isCellSelected(rowIndex: number, colId: string): boolean {
    if (!this.selectedRange) return false;

    const currentSheet = this.sheets[this.activeSheetIndex];
    const colIndex = currentSheet.columnHeaders.indexOf(colId);
    const startColIndex = currentSheet.columnHeaders.indexOf(this.selectedRange.startCol);
    const endColIndex = currentSheet.columnHeaders.indexOf(this.selectedRange.endCol);

    return (
      rowIndex >= this.selectedRange.startRow &&
      rowIndex <= this.selectedRange.endRow &&
      colIndex >= Math.min(startColIndex, endColIndex) &&
      colIndex <= Math.max(startColIndex, endColIndex)
    );
  }

  onCellContextMenu(event: MouseEvent, rowIndex: number, colId: string) {
    event.preventDefault();
    this.contextMenuPosition = { x: event.clientX, y: event.clientY };
    this.showContextMenu = true;
    this.selectedCell = { rowIndex, colId };
    this.selectedCells = {
      startRow: rowIndex,
      endRow: rowIndex,
      startCol: colId,
      endCol: colId
    };
  }

  onKeyDown(event: KeyboardEvent, rowIndex: number, colId: string) {
    if (event.key === 'Tab') {
      event.preventDefault();

      const currentSheet = this.sheets[this.activeSheetIndex];
      let nextColIndex = currentSheet.columnHeaders.indexOf(colId) + (event.shiftKey ? -1 : 1);
      let nextRowIndex = rowIndex;
      if (nextColIndex >= currentSheet.columnHeaders.length) {
        nextColIndex = 0;
        nextRowIndex++;
      } else if (nextColIndex < 0) {
        nextColIndex = currentSheet.columnHeaders.length - 1;
        nextRowIndex--;
      }
      if (nextRowIndex >= 0 && nextRowIndex < this.displayData.length) {
        const nextCell = document.querySelector(
          `td[data-row="${nextRowIndex}"][data-col="${currentSheet.columnHeaders[nextColIndex]}"]`
        ) as HTMLElement;

        if (nextCell) {
          nextCell.focus();
          this.onCellClick(
            { target: nextCell } as unknown as MouseEvent,
            nextRowIndex,
            currentSheet.columnHeaders[nextColIndex]
          );
        }
      }
    }
  }

  onCellMouseDown(event: MouseEvent, rowIndex: number, colId: string) {
    if (event.button === 0) {
      this.isSelecting = true;
      this.selectionStart = { row: rowIndex, col: colId };
      this.selectionEnd = { row: rowIndex, col: colId };
      this.selectedCell = { rowIndex, colId };
      this.updateSelectedRange();
    }
  }

  onCellMouseOver(event: MouseEvent, rowIndex: number, colId: string) {
    if (this.isSelecting && this.selectionStart) {
      this.selectionEnd = { row: rowIndex, col: colId };
      this.updateSelectedRange();
    }
  }

  onMouseLeave() {
    this.isSelecting = false;
  }

  onCellDragStart(event: DragEvent, rowIndex: number, colId: string) {
    if (!this.selectedRange || !this.isCellSelected(rowIndex, colId)) {
      return;
    }
    this.isDraggingCells = true;
    this.dragStartCell = { row: rowIndex, col: colId };
    this.draggedCellsData = this.getSelectedCellsData();
    if (event.dataTransfer) {
      event.dataTransfer.setData('text/plain', 'cells');
      event.dataTransfer.effectAllowed = 'move';
    }
  }

  onCellDragOver(event: DragEvent, rowIndex: number, colId: string) {
    if (this.isDraggingCells) {
      event.preventDefault();
      event.stopPropagation();
    }
  }

  onCellDrop(event: DragEvent, targetRow: number, targetCol: string) {
    event.preventDefault();
    event.stopPropagation();

    if (!this.isDraggingCells || !this.dragStartCell || !this.selectedRange) {
      return;
    }
    const rowOffset = targetRow - this.dragStartCell.row;
    const colOffset = this.sheets[this.activeSheetIndex].columnHeaders.indexOf(targetCol) -
      this.sheets[this.activeSheetIndex].columnHeaders.indexOf(this.dragStartCell.col);
    this.moveSelectedCells(rowOffset, colOffset);

    this.isDraggingCells = false;
    this.dragStartCell = null;
    this.draggedCellsData = [];
  }

  private getSelectedCellsData(): any[][] {
    if (!this.selectedRange) {
      if (this.selectedCell) {
        const currentSheet = this.sheets[this.activeSheetIndex];
        const colIndex = currentSheet.columnHeaders.indexOf(this.selectedCell.colId);
        return [[currentSheet.data[this.selectedCell.rowIndex][colIndex]]];
      }
      return [];
    }

    const data: any[][] = [];
    const currentSheet = this.sheets[this.activeSheetIndex];
    const startColIndex = currentSheet.columnHeaders.indexOf(this.selectedRange.startCol);
    const endColIndex = currentSheet.columnHeaders.indexOf(this.selectedRange.endCol);
    const minRow = Math.min(this.selectedRange.startRow, this.selectedRange.endRow);
    const maxRow = Math.max(this.selectedRange.startRow, this.selectedRange.endRow);
    const minCol = Math.min(startColIndex, endColIndex);
    const maxCol = Math.max(startColIndex, endColIndex);

    for (let i = minRow; i <= maxRow; i++) {
      const row: any[] = [];
      for (let j = minCol; j <= maxCol; j++) {
        row.push(currentSheet.data[i][j]);
      }
      data.push(row);
    }

    return data;
  }

  private moveSelectedCells(rowOffset: number, colOffset: number) {
    if (!this.selectedRange || !this.draggedCellsData.length) return;

    const currentSheet = this.sheets[this.activeSheetIndex];
    const newData = JSON.parse(JSON.stringify(currentSheet.data));
    for (let i = this.selectedRange.startRow; i <= this.selectedRange.endRow; i++) {
      for (let j = currentSheet.columnHeaders.indexOf(this.selectedRange.startCol);
        j <= currentSheet.columnHeaders.indexOf(this.selectedRange.endCol); j++) {
        newData[i][j] = '';
      }
    }
    this.draggedCellsData.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        const newRow = this.selectedRange!.startRow + rowOffset + rowIndex;
        const newCol = currentSheet.columnHeaders.indexOf(this.selectedRange!.startCol) + colOffset + colIndex;

        if (newRow >= 0 && newRow < newData.length &&
          newCol >= 0 && newCol < currentSheet.columnHeaders.length) {
          newData[newRow][newCol] = cell;
        }
      });
    });
    this.sheets[this.activeSheetIndex] = {
      ...currentSheet,
      data: newData
    };

    this.updateDisplayData();
  }

  isInDraggedSelection(rowIndex: number, colId: string): boolean {
    return this.isDraggingCells && this.isCellSelected(rowIndex, colId);
  }
  handleMenuAction(action: string) {
    if (!this.selectedCell) return;

    switch (action) {
      case 'cut':
        this.handleCut();
        break;
      case 'copy':
        this.handleCopy();
        break;
      case 'paste':
        this.handlePaste(this.selectedCell.rowIndex, this.selectedCell.colId);
        break;
      case 'insertRowAbove':
        this.insertRowAbove(this.selectedCell.rowIndex);
        break;
      case 'insertRowBelow':
        this.insertRowBelow(this.selectedCell.rowIndex);
        break;
      case 'insertColumnLeft':
        this.insertColumnLeft(this.selectedCell.colId);
        break;
      case 'insertColumnRight':
        this.insertColumnRight(this.selectedCell.colId);
        break;
      case 'deleteCurrentRow':
        this.deleteRow(this.selectedCell.rowIndex);
        break;
      case 'deleteRowAbove':
        this.deleteRowAbove(this.selectedCell.rowIndex);
        break;
      case 'deleteRowBelow':
        this.deleteRowBelow(this.selectedCell.rowIndex);
        break;
      case 'deleteCurrentColumn':
        this.deleteColumn(this.selectedCell.colId);
        break;
      case 'deleteColumnLeft':
        this.deleteColumnLeft(this.selectedCell.colId);
        this.clearSelection();
        break;
      case 'deleteColumnRight':
        this.deleteColumnRight(this.selectedCell.colId);
        this.clearSelection();
        break;
    }
    this.showContextMenu = false;
  }



  private async handleCut() {
    try {
      const currentSheet = this.sheets[this.activeSheetIndex];
      let startRow: number, endRow: number, startColIndex: number, endColIndex: number;
      if (!this.selectedRange && this.selectedCell) {
        startRow = endRow = this.selectedCell.rowIndex;
        startColIndex = endColIndex = currentSheet.columnHeaders.indexOf(this.selectedCell.colId);
        const cellValue = currentSheet.data[startRow][startColIndex];
        this.clipboardData = {
          data: [[cellValue]],
          startCol: this.selectedCell.colId,
          endCol: this.selectedCell.colId,
          startRow: startRow,
          endRow: endRow
        };
        currentSheet.data[startRow][startColIndex] = '';
        this.hf.setCellContents(
          {
            sheet: this.getActiveSheetId(),
            row: startRow,
            col: startColIndex
          },
          [['']]
        );
        await navigator.clipboard.writeText(cellValue?.toString() || '');
      }

      else if (this.selectedRange) {
        startRow = Math.min(this.selectedRange.startRow, this.selectedRange.endRow);
        endRow = Math.max(this.selectedRange.startRow, this.selectedRange.endRow);
        startColIndex = currentSheet.columnHeaders.indexOf(this.selectedRange.startCol);
        endColIndex = currentSheet.columnHeaders.indexOf(this.selectedRange.endCol);
        const selectedData: string[][] = [];
        for (let row = startRow; row <= endRow; row++) {
          const rowData: string[] = [];
          for (let col = startColIndex; col <= endColIndex; col++) {
            rowData.push(currentSheet.data[row][col]);
          }
          selectedData.push(rowData);
        }

        this.clipboardData = {
          data: selectedData,
          startCol: this.selectedRange.startCol,
          endCol: this.selectedRange.endCol,
          startRow: startRow,
          endRow: endRow
        };
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColIndex; col <= endColIndex; col++) {
            currentSheet.data[row][col] = '';
            this.hf.setCellContents(
              {
                sheet: this.getActiveSheetId(),
                row,
                col
              },
              [['']]
            );
          }
        }
        const textData = selectedData.map(row => row.join('\t')).join('\n');
        await navigator.clipboard.writeText(textData);
      }

      this.updateDisplayData();
    } catch (error) {
      console.error('Cut operation failed:', error);
    }
  }

  private getActiveSheetId(): number {
    return this.sheets[this.activeSheetIndex]?.id ?? 0;
  }

  private async handleCopy() {
    if (!this.selectedRange) return;

    try {
      this.clipboardData = {
        data: this.getSelectedCellsData(),
        startCol: this.selectedRange.startCol,
        endCol: this.selectedRange.endCol,
        startRow: this.selectedRange.startRow,
        endRow: this.selectedRange.endRow
      };

      const textData = this.clipboardData.data
        .map(row => row.join('\t'))
        .join('\n');
      await navigator.clipboard.writeText(textData);
    } catch (error) {
      console.error('Copy operation failed:', error);
    }
  }

  private async handlePaste(targetRow: number, targetCol: string) {
    try {
      const currentSheet = this.sheets[this.activeSheetIndex];
      let pasteData: string[][] = [];

      if (this.clipboardData) {
        pasteData = this.clipboardData.data;
      } else {
        const text = await navigator.clipboard.readText();
        if (text.includes('\t') || text.includes('\n')) {
          pasteData = text.split('\n').map(row => row.split('\t'));
        } else {
          pasteData = [[text]];
        }
      }

      const targetColIndex = currentSheet.columnHeaders.indexOf(targetCol);
      const newData = JSON.parse(JSON.stringify(currentSheet.data));
      pasteData.forEach((row, rowOffset) => {
        row.forEach((cell, colOffset) => {
          const pasteRowIndex = targetRow + rowOffset;
          const pasteColIndex = targetColIndex + colOffset;
          if (pasteRowIndex < currentSheet.data.length &&
            pasteColIndex < currentSheet.columnHeaders.length) {
            newData[pasteRowIndex][pasteColIndex] = cell?.trim() || '';
            this.hf.setCellContents(
              {
                sheet: this.getActiveSheetId(),
                row: pasteRowIndex,
                col: pasteColIndex
              },
              [[cell?.trim() || '']]
            );
          }
        });
      });
      this.sheets[this.activeSheetIndex] = {
        ...currentSheet,
        data: newData
      };

      this.updateDisplayData();
    } catch (error) {
      console.error('Paste operation failed:', error);
    }
  }

  private insertRowAbove(rowIndex: number) {
    const currentSheet = this.sheets[this.activeSheetIndex];
    const newRow = new Array(currentSheet.columnHeaders.length).fill('');
    currentSheet.data.splice(rowIndex, 0, newRow);
    if (currentSheet.data.length > 100) {
      currentSheet.data.pop();
    }
    this.updateDisplayData();
  }

  private insertRowBelow(rowIndex: number) {
    this.insertRowAbove(rowIndex + 1);
  }

  private insertColumnLeft(colId: string) {
    const currentSheet = this.sheets[this.activeSheetIndex];
    const colIndex = currentSheet.columnHeaders.indexOf(colId);

    if (colIndex === -1) return;
    const oldData = JSON.parse(JSON.stringify(currentSheet.data));
    const newData = currentSheet.data.map((row, rowIndex) => {
      const newRow = [...row];
      newRow.splice(colIndex, 0, '');
      for (let i = colIndex + 1; i < newRow.length; i++) {
        newRow[i] = oldData[rowIndex][i - 1];
      }
      return newRow;
    });

    const newHeaders = [...currentSheet.columnHeaders];
    newHeaders.splice(colIndex, 0, 'TEMP');
    this.sheets[this.activeSheetIndex] = {
      ...currentSheet,
      data: newData,
      columnHeaders: newHeaders
    };
    this.updateColumnNames();
    this.updateDisplayData();
  }


  private deleteColumn(colId: string) {
    const currentSheet = this.sheets[this.activeSheetIndex];
    const colIndex = currentSheet.columnHeaders.indexOf(colId);
    currentSheet.data.forEach(row => {
      row.splice(colIndex, 1);
    });
    currentSheet.columnHeaders.splice(colIndex, 1);
    this.updateColumnNames();
    this.updateDisplayData();
  }

  private insertColumnRight(colId: string) {
    const currentSheet = this.sheets[this.activeSheetIndex];
    const colIndex = currentSheet.columnHeaders.indexOf(colId);

    if (colIndex === -1) return;
    const insertIndex = colIndex + 1;
    const newData = currentSheet.data.map(row => {
      const newRow = [...row];
      newRow.splice(insertIndex, 0, '');
      return newRow;
    });

    const newHeaders = [...currentSheet.columnHeaders];
    newHeaders.splice(insertIndex, 0, 'TEMP');
    this.sheets[this.activeSheetIndex] = {
      ...currentSheet,
      data: newData,
      columnHeaders: newHeaders
    };
    this.updateColumnNames();
    this.updateDisplayData();
  }


  private deleteRow(rowIndex: number) {
    const currentSheet = this.sheets[this.activeSheetIndex];
    currentSheet.data.splice(rowIndex, 1);
    currentSheet.data.push(new Array(currentSheet.columnHeaders.length).fill(''));
    this.updateDisplayData();
  }

  selectSheet(index: number): void {
    if (index >= 0 && index < this.sheets.length && index !== this.activeSheetIndex) {
      this.activeSheetIndex = index;
      this.updateDisplayData();
      this.selectedCell = null;
      this.selectedRange = null;
    }
  }

  private deleteColumnLeft(colId: string) {
    const currentSheet = this.sheets[this.activeSheetIndex];
    const currentColIndex = currentSheet.columnHeaders.indexOf(colId);
    if (currentColIndex > 0) {
      const targetColIndex = currentColIndex - 1;
      const oldHeaders = [...currentSheet.columnHeaders];
      const newData = currentSheet.data.map(row => {
        const newRow = [...row];
        newRow.splice(targetColIndex, 1);
        return newRow;
      });
      const newHeaders = [...currentSheet.columnHeaders];
      newHeaders.splice(targetColIndex, 1);
      this.sheets[this.activeSheetIndex] = {
        ...currentSheet,
        data: newData,
        columnHeaders: newHeaders
      };
      this.updateColumnNames();
      this.updateDisplayData();
      if (this.selectedCell) {
        const selectedColIndex = oldHeaders.indexOf(this.selectedCell.colId);
        if (selectedColIndex > targetColIndex) {
          this.selectedCell.colId = this.sheets[this.activeSheetIndex].columnHeaders[selectedColIndex - 1];
        }
      }
    }
  }

  private deleteColumnRight(colId: string) {
    const currentSheet = this.sheets[this.activeSheetIndex];
    const currentColIndex = currentSheet.columnHeaders.indexOf(colId);
    if (currentColIndex < currentSheet.columnHeaders.length - 1) {
      const targetColIndex = currentColIndex + 1;
      const oldHeaders = [...currentSheet.columnHeaders];
      const newData = currentSheet.data.map(row => {
        const newRow = [...row];
        newRow.splice(targetColIndex, 1);
        return newRow;
      });
      const newHeaders = [...currentSheet.columnHeaders];
      newHeaders.splice(targetColIndex, 1);
      this.sheets[this.activeSheetIndex] = {
        ...currentSheet,
        data: newData,
        columnHeaders: newHeaders
      };
      this.updateColumnNames();
      this.updateDisplayData();
    }
  }

  private deleteRowAbove(rowIndex: number) {
    const currentSheet = this.sheets[this.activeSheetIndex];

    if (rowIndex > 0) {
      currentSheet.data.splice(rowIndex - 1, 1);
      currentSheet.data.push(new Array(currentSheet.columnHeaders.length).fill(''));
      this.updateDisplayData();
    }
  }

  private deleteRowBelow(rowIndex: number) {
    const currentSheet = this.sheets[this.activeSheetIndex];

    if (rowIndex < currentSheet.data.length - 1) {
      currentSheet.data.splice(rowIndex + 1, 1);
      currentSheet.data.push(new Array(currentSheet.columnHeaders.length).fill(''));
      this.updateDisplayData();
    }
  }

  private formulaInput: HTMLInputElement | null = null;
  private isFormulaBarFocused = false;
  currentFormula = '';

  // Initialize the reference in ngOnInit
  ngOnInit() {
    // Keep your existing ngOnInit code
    this.selectSheet(0);
    this.formulaInput = document.getElementById('formula-input') as HTMLInputElement;
    if (this.formulaInput) {
      this.formulaInput.value = '';
    }

    // Add this line to get the grid container reference
    this.gridContainerRef = document.querySelector('.grid-container');
  }

  onCellClick(event: MouseEvent, rowIndex: number, colId: string) {
    const currentSheet = this.sheets[this.activeSheetIndex];
    const colIndex = currentSheet.columnHeaders.indexOf(colId);
    const cellValue = currentSheet.data[rowIndex][colIndex] || '';

    this.selectedCell = { rowIndex, colId };

    // Update formula bar
    if (this.formulaInput) {
      this.formulaInput.value = cellValue;
      this.currentFormula = cellValue;
    }

    // If it's a formula cell, prevent it from becoming editable
    if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
      const cell = event.target as HTMLElement;
      cell.contentEditable = 'false';
      setTimeout(() => {
        cell.contentEditable = 'true';
      }, 100);
    }

    this.selectionStart = { row: rowIndex, col: colId };
    this.selectionEnd = { row: rowIndex, col: colId };
    this.updateSelectedRange();
  }

  private updateSelectedRange() {
    if (!this.selectionStart || !this.selectionEnd) return;

    const currentSheet = this.sheets[this.activeSheetIndex];
    const startColIndex = currentSheet.columnHeaders.indexOf(this.selectionStart.col);
    const endColIndex = currentSheet.columnHeaders.indexOf(this.selectionEnd.col);

    this.selectedRange = {
      startRow: Math.min(this.selectionStart.row, this.selectionEnd.row),
      endRow: Math.max(this.selectionStart.row, this.selectionEnd.row),
      startCol: currentSheet.columnHeaders[Math.min(startColIndex, endColIndex)],
      endCol: currentSheet.columnHeaders[Math.max(startColIndex, endColIndex)]
    };
    if (this.formulaInput && this.selectedRange && !this.isDirectCellEdit) {
      if (this.selectedRange.startRow === this.selectedRange.endRow &&
        this.selectedRange.startCol === this.selectedRange.endCol) {
        const value = this.getCellValue(this.selectedRange.startRow, this.selectedRange.startCol);
        this.formulaInput.value = value;
        this.currentFormula = value;
      } else if (this.isFormulaBarFocused) {
        const rangeText = `${this.selectedRange.startCol}${this.selectedRange.startRow + 1}:${this.selectedRange.endCol}${this.selectedRange.endRow + 1}`;
        const cursorPos = this.formulaInput.selectionStart || 0;
        const currentValue = this.formulaInput.value;
        this.formulaInput.value =
          currentValue.slice(0, cursorPos) +
          rangeText +
          currentValue.slice(cursorPos);
        this.currentFormula = this.formulaInput.value;
      }
    }
  }

  // onFormulaBarInput(event: Event) {
  //   const input = event.target as HTMLInputElement;
  //   this.currentFormula = input.value;
  // }

  // onFormulaEnter(event: KeyboardEvent) {
  //   event.preventDefault();
  //   const formula = this.currentFormula;
  //   const currentSheet = this.sheets[this.activeSheetIndex];

  //   // Parse the formula with potential cell reference
  //   const { targetCell, formula: processedFormula } = FormulaParser.parseFormulaWithCellRef(
  //     formula,
  //     currentSheet.data,
  //     currentSheet.columnHeaders
  //   );

  //   if (!processedFormula) {
  //     console.error('Invalid formula format');
  //     return;
  //   }

  //   if (targetCell) {
  //     // Case: A1=SUM(B1:B8)
  //     const result = FormulaParser.evaluateFormula(processedFormula, currentSheet.data, currentSheet.columnHeaders);
  //     this.updateCellValue(targetCell.row, targetCell.col, processedFormula);

  //     // Update the cell display
  //     const cell = document.querySelector(
  //       `td[data-row="${targetCell.row}"][data-col="${targetCell.col}"]`
  //     ) as HTMLElement;

  //     if (cell) {
  //       cell.textContent = result.toString();
  //     }
  //   } else if (this.selectedCell) {
  //     // Case: Regular formula without cell reference
  //     const { rowIndex, colId } = this.selectedCell;
  //     this.updateCellValue(rowIndex, colId, processedFormula);
  //   }

  //   this.isFormulaBarFocused = false;

  //   // Clear the formula bar if needed
  //   if (this.formulaInput) {
  //     this.formulaInput.value = '';
  //     this.currentFormula = '';
  //   }
  // }

  // private updateCellValue(rowIndex: number, colId: string, value: string) {
  //   const currentSheet = this.sheets[this.activeSheetIndex];
  //   const colIndex = currentSheet.columnHeaders.indexOf(colId);

  //   // Store the formula in the cell
  //   currentSheet.data[rowIndex][colIndex] = value;

  //   if (value.startsWith('=')) {
  //     try {
  //       const result = FormulaParser.evaluateFormula(
  //         value,
  //         currentSheet.data,
  //         currentSheet.columnHeaders
  //       );

  //       const cell = document.querySelector(
  //         `td[data-row="${rowIndex}"][data-col="${colId}"]`
  //       ) as HTMLElement;

  //       if (cell) {
  //         cell.textContent = result.toString();
  //       }
  //     } catch (error) {
  //       console.error('Formula evaluation error:', error);
  //       const cell = document.querySelector(
  //         `td[data-row="${rowIndex}"][data-col="${colId}"]`
  //       ) as HTMLElement;
  //       if (cell) {
  //         cell.textContent = '#ERROR!';
  //       }
  //     }
  //   }

  //   this.updateDisplayData();
  // }


  @HostListener('window:keydown', ['$event'])
  handleKeyboardShortcut(event: KeyboardEvent) {

    console.log('Key pressed:', event.key);
    console.log('Is formula bar focused:', this.isFormulaBarFocused);

    if (this.isFormulaBarFocused) {
      if (event.key === 'Enter') {
        console.log("Enter is pressed in formula bar");
        this.onFormulaEnter(event);
      }
      // Prevent any other keyboard shortcuts while formula bar is focused
      return;
    }

    if (event.key.startsWith('Arrow') && this.selectedCell) {
      event.preventDefault();
      this.navigateWithArrowKeys(event.key);
      return;
    }

    if (event.ctrlKey || event.metaKey) {
      switch (event.key.toLowerCase()) {
        case 'c':
          event.preventDefault();
          this.handleCopy();
          break;
        case 'v':
          event.preventDefault();
          this.handlePaste(
            this.selectedCell?.rowIndex || 0,
            this.selectedCell?.colId || 'A'
          );
          break;
        case 'x':
          event.preventDefault();
          this.handleCut();
          break;
      }
    }
  }

  private navigateWithArrowKeys(key: string) {
    if (!this.selectedCell) return;

    const currentSheet = this.sheets[this.activeSheetIndex];
    const { rowIndex, colId } = this.selectedCell;
    const currentColIndex = currentSheet.columnHeaders.indexOf(colId);
    let nextRow = rowIndex;
    let nextColIndex = currentColIndex;

    switch (key) {
      case 'ArrowUp':
        nextRow = Math.max(0, rowIndex - 1);
        break;
      case 'ArrowDown':
        nextRow = Math.min(currentSheet.data.length - 1, rowIndex + 1);
        break;
      case 'ArrowLeft':
        nextColIndex = Math.max(0, currentColIndex - 1);
        break;
      case 'ArrowRight':
        nextColIndex = Math.min(currentSheet.columnHeaders.length - 1, currentColIndex + 1);
        break;
    }

    if (nextRow !== rowIndex || nextColIndex !== currentColIndex) {
      const nextColId = currentSheet.columnHeaders[nextColIndex];

      const nextCell = document.querySelector(
        `td[data-row="${nextRow}"][data-col="${nextColId}"]`
      ) as HTMLElement;

      if (nextCell) {
        nextCell.focus();
        this.onCellClick(
          { target: nextCell } as unknown as MouseEvent,
          nextRow,
          nextColId
        );
      }
    }
  }

  sortState: SortState = {
    column: '',
    direction: null,
    lastClickTime: 0
  };

  private readonly DOUBLE_CLICK_THRESHOLD = 300;

  private sortColumn(column: string, direction: 'asc' | 'desc') {
    const currentSheet = this.sheets[this.activeSheetIndex];
    const colIndex = currentSheet.columnHeaders.indexOf(column);
    const indexedData = currentSheet.data.map((row, index) => ({
      index,
      value: this.getCellSortValue(row[colIndex])
    }));

    indexedData.sort((a, b) => {
      if (a.value === '' && b.value === '') return 0;
      if (a.value === '') return direction === 'asc' ? 1 : -1;
      if (b.value === '') return direction === 'asc' ? -1 : 1;

      const comparison = this.compareValues(a.value, b.value);
      return direction === 'asc' ? comparison : -comparison;
    });

    const newData = indexedData.map(({ index }) => [...currentSheet.data[index]]);

    this.sheets[this.activeSheetIndex] = {
      ...currentSheet,
      data: newData
    };

    this.updateDisplayData();
  }

  private getCellSortValue(value: any): any {
    if (value === null || value === undefined || value === '') {
      return '';
    }

    // Handle formulas - get the calculated value
    if (typeof value === 'string' && value.startsWith('=')) {
      try {
        const result = FormulaParser.parseFormula(
          value,
          this.sheets[this.activeSheetIndex].data,
          this.sheets[this.activeSheetIndex].columnHeaders
        );
        return result;
      } catch (error) {
        return value;
      }
    }

    // Handle numbers stored as strings
    if (typeof value === 'string' && !isNaN(Number(value))) {
      return Number(value);
    }

    return value;
  }

  private compareValues(a: any, b: any): number {
    if (typeof a !== typeof b) {
      return String(a).localeCompare(String(b));
    }
    if (typeof a === 'number' && typeof b === 'number') {
      return a - b;
    }
    if (typeof a === 'string' && typeof b === 'string') {
      return a.toLowerCase().localeCompare(b.toLowerCase());
    }
    return 0;
  }

  handleColumnHeaderClick(event: MouseEvent, column: string) {
    const currentTime = new Date().getTime();
    const timeDiff = currentTime - this.sortState.lastClickTime;
    if (this.sortState.column === column && timeDiff < this.DOUBLE_CLICK_THRESHOLD) {
      if (this.sortState.direction === 'asc') {
        this.sortState.direction = 'desc';
      } else if (this.sortState.direction === 'desc') {
        this.sortState.direction = null;
        this.sortState.column = '';
        this.updateDisplayData();
        this.sortState.lastClickTime = 0;
        return;
      }
    } else {
      this.sortState.column = column;
      this.sortState.direction = 'asc';
    }

    this.sortState.lastClickTime = currentTime;

    if (this.sortState.direction) {
      this.sortColumn(column, this.sortState.direction);
    }
  }

  onFormulaBarFocus() {
    this.isFormulaBarFocused = true;
  }

  onFormulaBarBlur() {
    this.isFormulaBarFocused = false;
  }


  // Helper method to get cell key
  private getCellKey(row: number, col: string): string {
    return `${row}-${col}`;
  }

  // Update cell metadata
  setCellType(row: number, col: string, type: 'text' | 'number' | 'date') {
    const key = this.getCellKey(row, col);
    const existing = this.cellMetadata.get(key) || { type: 'text' };
    this.cellMetadata.set(key, { ...existing, type });
  }

  // Update onCellEdit method
  onCellEdit(event: Event, rowIndex: number, colId: string) {
    const target = event.target as HTMLElement;
    const value = target.textContent?.trim() || '';
    const key = this.getCellKey(rowIndex, colId);
    const metadata = this.cellMetadata.get(key) || { type: 'text' };

    // Validate input
    const validation = DataValidator.validateCell(value, metadata.type);
    if (!validation.isValid) {
      target.classList.add('invalid-input');
      target.title = validation.message || 'Invalid input';
      return;
    }

    target.classList.remove('invalid-input');
    target.title = '';
    this.updateCellValue(rowIndex, colId, value);
  }

  // Enhanced formula parsing with absolute references
  private parseFormulaReference(ref: string): { row: number; col: string; isAbsolute: { row: boolean; col: boolean } } {
    const match = ref.match(/(\$?)([A-Z]+)(\$?)(\d+)/);
    if (!match) throw new Error('Invalid cell reference');

    return {
      col: match[2],
      row: parseInt(match[4]) - 1,
      isAbsolute: {
        col: match[1] === '$',
        row: match[3] === '$'
      }
    };
  }

  // Add method to adjust relative references when copying formulas
  private adjustFormulaReferences(formula: string, rowOffset: number, colOffset: number): string {
    return formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (match, colAbs, col, rowAbs, row) => {
      if (colAbs === '$' && rowAbs === '$') return match;

      let newCol = col;
      let newRow = parseInt(row);

      if (!colAbs) {
        // Adjust column reference
        newCol = this.offsetColumn(col, colOffset);
      }

      if (!rowAbs) {
        // Adjust row reference
        newRow += rowOffset;
      }

      return `${colAbs}${newCol}${rowAbs}${newRow}`;
    });
  }

  // Helper method to offset column letters
  private offsetColumn(col: string, offset: number): string {
    const base = 'A'.charCodeAt(0);
    const current = col.split('').reduce((acc, char) =>
      acc * 26 + char.charCodeAt(0) - base + 1, 0);
    const newVal = current + offset;

    let result = '';
    let value = newVal;
    while (value > 0) {
      value--;
      result = String.fromCharCode(base + (value % 26)) + result;
      value = Math.floor(value / 26);
    }
    return result;
  }

  private isSpacePressed = false;
  private gridContainerRef: HTMLElement | null = null;

  // Add these new event listeners
  @HostListener('window:keydown', ['$event'])
  handleKeyDown(event: KeyboardEvent) {
    if (event.code === 'Space' && !this.isSpacePressed) {
      event.preventDefault(); // Prevent space from scrolling page
      this.isSpacePressed = true;
    }
  }

  @HostListener('window:keyup', ['$event'])
  handleKeyUp(event: KeyboardEvent) {
    if (event.code === 'Space') {
      this.isSpacePressed = false;
    }
  }

  @HostListener('wheel', ['$event'])
  handleWheel(event: WheelEvent) {
    if (this.isSpacePressed) {
      event.preventDefault();

      // Get the grid container
      if (!this.gridContainerRef) {
        this.gridContainerRef = document.querySelector('.grid-container');
      }

      if (this.gridContainerRef) {
        // Scroll horizontally instead of vertically
        this.gridContainerRef.scrollLeft += event.deltaY;
      }
    }
  }



  // Add a cleanup method
  ngOnDestroy() {
    this.gridContainerRef = null;
  }

  // =====================================================

  handleFileAction(action: { action: string, data?: any }) {
    switch (action.action) {
      case 'new':
        this.createNewFile();
        break;
      case 'open':
        this.openFile(action.data);
        break;
    }
  }

  createNewFile() {
    // Clear HyperFormula instance
    this.hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });

    // Create new sheet with default data
    const defaultColumnHeaders = this.generateColumnHeaders();
    const newSheetId = this.hf.addSheet('Sheet 1');

    const newSheet = {
      name: 'Sheet 1',
      data: this.generateInitialData(),
      id: typeof newSheetId === 'number' ? newSheetId : 0,
      columnHeaders: [...defaultColumnHeaders]
    };

    // Reset state
    this.sheets = [newSheet];
    this.activeSheetIndex = 0;
    this.selectedRange = null;
    this.selectedCell = null;
    this.currentFormula = '';
    this.cellMetadata.clear();
    this.cellFormatting.clear();

    // Update display
    this.updateDisplayData();
    this.saveToStorage();
  }

  private openFile(sheets: any[]) {
    if (!sheets || !sheets.length) return;

    // Clear existing sheets
    this.sheets.forEach(sheet => {
      try {
        this.hf.removeSheet(sheet.id);
      } catch (error) {
        console.error('Error removing sheet:', error);
      }
    });

    // Add new sheets
    this.sheets = sheets.map(sheet => {
      const sheetId = this.hf.addSheet(sheet.name);
      return {
        ...sheet,
        id: typeof sheetId === 'number' ? sheetId : 0
      };
    });

    this.activeSheetIndex = 0;
    this.selectedRange = null;
    this.selectedCell = null;
    this.currentFormula = '';
    this.updateDisplayData();
  }


  // @HostListener('window:keydown', ['$event'])
  // handleKeyboardShortcut(event: KeyboardEvent) {
  //   // Log the current state for debugging
  //   console.log('Key pressed:', event.key);
  //   console.log('Formula bar focused:', this.isFormulaBarFocused);
  //   console.log('Current formula:', this.currentFormula);

  //   // Check if we're in the formula bar
  //   const activeElement = document.activeElement;
  //   const isFormulaInput = activeElement?.id === 'formula-input';

  //   if (isFormulaInput && event.key === 'Enter') {
  //     console.log("Enter pressed in formula bar");
  //     event.preventDefault(); // Prevent default Enter behavior
  //     this.onFormulaEnter(event);
  //   }
  // }

  onFormulaBarInput(event: Event) {
    const input = event.target as HTMLInputElement;
    this.currentFormula = input.value;
    console.log('Formula input changed:', this.currentFormula);
  }

  onFormulaEnter(event: KeyboardEvent) {
    console.log('Processing formula:', this.currentFormula);

    const formula = this.currentFormula;
    const currentSheet = this.sheets[this.activeSheetIndex];

    // Parse the formula with potential cell reference
    const { targetCell, formula: processedFormula } = FormulaParser.parseFormulaWithCellRef(
      formula,
      currentSheet.data,
      currentSheet.columnHeaders
    );

    if (!processedFormula) {
      console.error('Invalid formula format');
      return;
    }

    if (targetCell) {
      // Case: A1=SUM(B1:B8)
      console.log('Target cell:', targetCell);
      const result = FormulaParser.evaluateFormula(
        processedFormula,
        currentSheet.data,
        currentSheet.columnHeaders
      );
      console.log('Formula result:', result);

      this.updateCellValue(targetCell.row, targetCell.col, processedFormula);
    } else if (this.selectedCell) {
      // Case: Regular formula
      const { rowIndex, colId } = this.selectedCell;
      this.updateCellValue(rowIndex, colId, processedFormula);
    }

    // Update UI and state
    this.isFormulaBarFocused = false;
    this.updateDisplayData();
  }

  private updateCellValue(rowIndex: number, colId: string, value: string) {
    console.log('Updating cell value:', { rowIndex, colId, value });

    const currentSheet = this.sheets[this.activeSheetIndex];
    const colIndex = currentSheet.columnHeaders.indexOf(colId);

    // Store the formula in the cell
    currentSheet.data[rowIndex][colIndex] = value;

    if (value.startsWith('=')) {
      try {
        const result = FormulaParser.evaluateFormula(
          value,
          currentSheet.data,
          currentSheet.columnHeaders
        );

        console.log('Evaluated result:', result);

        const cell = document.querySelector(
          `td[data-row="${rowIndex}"][data-col="${colId}"]`
        ) as HTMLElement;

        if (cell) {
          cell.textContent = result.toString();
        }
      } catch (error) {
        console.error('Formula evaluation error:', error);
        const cell = document.querySelector(
          `td[data-row="${rowIndex}"][data-col="${colId}"]`
        ) as HTMLElement;
        if (cell) {
          cell.textContent = '#ERROR!';
        }
      }
    }

    this.updateDisplayData();
  }


}
