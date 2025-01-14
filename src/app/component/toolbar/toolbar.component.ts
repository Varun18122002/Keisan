import { Component, EventEmitter, Input, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import { SaveDialogComponent } from '../utils/save-dialog/save-dialog.component';

interface SaveDetails {
  filename: string;
  format: 'xlsx' | 'csv';
}

@Component({
  selector: 'app-toolbar',
  standalone: true,
  imports: [CommonModule, SaveDialogComponent],
  template: `
    <div class="toolbar">
      <div class="toolbar-container">
        <div class="dropdown">
          <button class="dropdown-btn">File</button>
          <div class="dropdown-content">
            <button (click)="onNewSheet()">New Sheet</button>
            <button (click)="onOpenSheet()">Open Sheet</button>
            <button (click)="onSaveSheet()">Save</button>

          </div>
        </div>


        <div class="dropdown">
          <button class="dropdown-btn">Edit</button>
          <div class="dropdown-content">
            <button (click)="onEditAction('cut')">Cut</button>
            <button (click)="onEditAction('copy')">Copy</button>
            <button (click)="onEditAction('paste')">Paste</button>
          <hr/>
            <button (click)="onEditAction('insertRowAbove')">Insert Row Above</button>
            <button (click)="onEditAction('insertRowBelow')">Insert Row Below</button>
            <button (click)="onEditAction('insertColumnLeft')">Insert Column Left</button>
            <button (click)="onEditAction('insertColumnRight')">Insert Column Right</button>
          <hr/>
            <button (click)="onEditAction('deleteCurrentRow')">Delete Row</button>
            <button (click)="onEditAction('deleteCurrentColumn')">Delete Column</button>
        </div>
        </div>

        <div class="formatting-tools">
      <button
            [class.active]="formatting.bold"
            (click)="onFormatting('bold')"
      >
        <strong>B</strong>
      </button>
      <button
            [class.active]="formatting.italic"
            (click)="onFormatting('italic')"
      >
        <em>I</em>
      </button>
      <button
            [class.active]="formatting.underline"
            (click)="onFormatting('underline')"
      >
        <u>U</u>
      </button>
      </div>
    </div>

      <!-- Formula Bar with Cell Reference -->
      <div class="formula-section">
        <div class="cell-reference" [class.active]="selectedRange">
          {{ getCellReferenceDisplay() }}
        </div>
        <div class="formula-bar">
        <label for="formula-input">fx:</label>
        <input
    #formulaInput
    id="formula-input"
    type="text"
    [value]="currentFormula"
    (input)="onFormulaInput($event)"
    (keydown)="handleKeyDown($event)"
    (focus)="formulaFocus.emit()"
    (blur)="formulaBlur.emit()"
    placeholder="Enter formula here"
  />
      </div>
        </div>

    </div>

    <!-- Confirmation Dialog -->
    <div *ngIf="showConfirmDialog" class="dialog-overlay">
      <div class="dialog-content">
        <h2>Save Changes?</h2>
        <p>Do you want to save changes before creating a new file?</p>
        <div class="dialog-buttons">
          <button (click)="handleConfirmDialogResponse('save')">Save</button>
          <button (click)="handleConfirmDialogResponse('dont-save')">Don't Save</button>
          <button (click)="handleConfirmDialogResponse('cancel')">Cancel</button>
        </div>
      </div>
    </div>

    <app-save-dialog
      *ngIf="showSaveDialog"
      (save)="handleSave($event)"
      (close)="showSaveDialog = false"
    />
  `,

  styleUrl: 'toolbar.component.scss'
})


export class ToolbarComponent {

  @Input() gridData: any[][] = [];
  @Input() formatting: any = {};
  @Input() selectedRange: any = null;
  @Input() currentFormula: string = '';

  @Output() formulaInput = new EventEmitter<Event>();
  @Output() formulaEnter = new EventEmitter<KeyboardEvent>();
  @Output() formulaFocus = new EventEmitter<void>();
  @Output() formulaBlur = new EventEmitter<void>();
  @Output() formatChange = new EventEmitter<{ type: string, value: boolean }>();
  @Output() editAction = new EventEmitter<string>();
  @Output() fileAction = new EventEmitter<{ action: string, data?: any }>();

  showSaveDialog = false;




  showConfirmDialog = false;
  private fileInput: HTMLInputElement;

  constructor() {
    this.fileInput = document.createElement('input');
    this.fileInput.type = 'file';
    this.fileInput.accept = '.xlsx,.csv';
    this.fileInput.addEventListener('change', this.handleFileSelect.bind(this));
  }

  getCellReferenceDisplay(): string {
    if (!this.selectedRange) return '';

    if (
      this.selectedRange.startRow === this.selectedRange.endRow &&
      this.selectedRange.startCol === this.selectedRange.endCol
    ) {
      return `${this.selectedRange.startCol}${this.selectedRange.startRow + 1}`;
    }

    return `${this.selectedRange.startCol}${this.selectedRange.startRow + 1}:${this.selectedRange.endCol}${this.selectedRange.endRow + 1}`;
  }

  onFormulaEnter(event: KeyboardEvent) {
    event.preventDefault(); // Prevent default enter behavior
    this.formulaEnter.emit(event);
  }

  onFormatting(type: 'bold' | 'italic' | 'underline') {
    this.formatting[type] = !this.formatting[type];
    this.formatChange.emit({ type, value: this.formatting[type] });
  }

  onEditAction(action: string) {
    this.editAction.emit(action);
  }
  onFormulaFocus() {
    this.formulaFocus.emit();
  }

  onFormulaBlur() {
    this.formulaBlur.emit();
  }





  onNewSheet() {
    // Check if there are unsaved changes
    if (this.hasUnsavedChanges()) {
      this.showConfirmDialog = true;
    } else {
      this.createNewFile();
    }
  }

  private hasUnsavedChanges(): boolean {
    // Implement your logic to check for unsaved changes
    // For example, compare current state with last saved state
    return true; // Default to true for safety
  }

  handleConfirmDialogResponse(response: 'save' | 'dont-save' | 'cancel') {
    this.showConfirmDialog = false;

    switch (response) {
      case 'save':
        this.showSaveDialog = true;
        // After save is complete, create new file
        break;
      case 'dont-save':
        this.createNewFile();
        break;
      case 'cancel':
        // Do nothing, just close dialog
        break;
    }
  }

  private createNewFile() {
    this.fileAction.emit({ action: 'new' });
  }

  onOpenSheet() {
    this.fileInput.click();
  }

  private handleFileSelect(event: Event) {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];

    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e: ProgressEvent<FileReader>) => {
      const data = e.target?.result;
      const workbook = XLSX.read(data, { type: 'binary' });

      // Convert workbook to your app's data format
      const sheets = workbook.SheetNames.map(name => {
        const sheet = workbook.Sheets[name];
        return {
          name,
          data: XLSX.utils.sheet_to_json(sheet, { header: 1 }),
          id: workbook.SheetNames.indexOf(name),
          columnHeaders: this.generateColumnHeaders(XLSX.utils.decode_range(sheet['!ref'] || 'A1').e.c + 1)
        };
      });

      this.fileAction.emit({ action: 'open', data: sheets });
    };

    reader.readAsBinaryString(file);
    input.value = ''; // Reset input for future uploads
  }

  private generateColumnHeaders(count: number): string[] {
    return Array.from({ length: count }, (_, i) =>
      String.fromCharCode(65 + i % 26).repeat(Math.floor(i / 26) + 1)
    );
  }

  onSaveSheet() {
    this.showSaveDialog = true;
  }

  handleSave(saveDetails: SaveDetails) {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(this.gridData);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    try {
      if (saveDetails.format === 'xlsx') {
        XLSX.writeFile(workbook, `${saveDetails.filename}.xlsx`);
      } else {
        const csv = XLSX.utils.sheet_to_csv(worksheet);
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `${saveDetails.filename}.csv`;
        link.click();
      }
      this.showSaveDialog = false;
    } catch (error) {
      console.error('Error saving file:', error);
    }
  }

  onFormulaInput(event: Event) {
    console.log('Toolbar formula input:', (event.target as HTMLInputElement).value);
    this.formulaInput.emit(event);
  }

  handleKeyDown(event: KeyboardEvent) {
    console.log('Toolbar key down:', event.key);
    if (event.key === 'Enter') {
      event.preventDefault();
      console.log('Toolbar Enter pressed');
      this.formulaEnter.emit(event);
    }
  }



}
