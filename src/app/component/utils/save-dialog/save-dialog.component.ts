import { Component, EventEmitter, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

export interface SaveDetails {
  filename: string;
  format: 'xlsx' | 'csv';
}


@Component({
  selector: 'app-save-dialog',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
    <div class="dialog-overlay" (click)="onClose()">
      <div class="dialog-content" (click)="$event.stopPropagation()">
        <h2 class="text-lg font-semibold mb-4">Save Spreadsheet</h2>

        <div class="mb-4">
          <label class="block mb-2">Filename:</label>
          <input
            type="text"
            [(ngModel)]="filename"
            class="w-full px-3 py-2 border rounded"
            placeholder="Enter filename"
          />
        </div>

        <div class="mb-4">
          <label class="block mb-2">Format:</label>
          <select
            [(ngModel)]="format"
            class="w-full px-3 py-2 border rounded"
          >
            <option value="xlsx">Excel (.xlsx)</option>
            <option value="csv">CSV (.csv)</option>
          </select>
        </div>

        <div class="flex justify-end gap-2">
          <button
            class="px-4 py-2 border rounded"
            (click)="onClose()"
          >
            Cancel
          </button>
          <button
            class="px-4 py-2 bg-blue-500 text-white rounded"
            (click)="onSave()"
            [disabled]="!filename"
          >
            Save
          </button>
        </div>
      </div>
    </div>
  `,
  styles: [`
    .dialog-overlay {
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(0, 0, 0, 0.5);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 1000;
    }

    .dialog-content {
      background: white;
      padding: 2rem;
      border-radius: 0.5rem;
      min-width: 400px;
    }
  `]
})


export class SaveDialogComponent {
  @Output() save = new EventEmitter<SaveDetails>();
  @Output() close = new EventEmitter<void>();

  filename = '';
  format: 'xlsx' | 'csv' = 'xlsx';

  onSave() {
    if (this.filename) {
      // Add extension if not present
      const fullFilename = this.filename.endsWith(`.${this.format}`)
        ? this.filename
        : `${this.filename}.${this.format}`;

      this.save.emit({
        filename: fullFilename,
        format: this.format
      });
    }
  }

  onClose() {
    this.close.emit();
  }
}
