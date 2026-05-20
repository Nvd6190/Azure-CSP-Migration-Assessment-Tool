import { Component, EventEmitter, Input, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { MigrationService } from '../services/migration.service';
import { AssessmentResponse } from '../models/assessment.model';

@Component({
  selector: 'app-upload',
  standalone: true,
  imports: [CommonModule],
  template: `
    <div class="upload-card">
      <div class="upload-header">
        <h2>{{ mode === 'aws' ? 'Upload AWS Resource Inventory' : mode === 'gcp' ? 'Upload GCP Resource Inventory' : 'Upload Azure Resource Export' }}</h2>
        <p>Upload your {{ mode === 'aws' ? 'AWS resource inventory (.xlsx or .json) to map services to Azure equivalents with SKU recommendations' : mode === 'gcp' ? 'GCP resource inventory (.xlsx or .csv) to map services to Azure equivalents with migration guidance' : mode === 'jio' ? 'Azure resource .xlsx export file to assess Jio region availability' : mode === 'region' ? 'Azure resource .xlsx export file to assess region move support' : 'Azure resource .xlsx export file to assess subscription move support' }}</p>
      </div>

      <div
        class="drop-zone"
        [class.drag-over]="isDragOver"
        [class.has-file]="!!selectedFile"
        (dragover)="onDragOver($event)"
        (dragleave)="onDragLeave($event)"
        (drop)="onDrop($event)"
        (click)="fileInput.click()">

        <input
          #fileInput
          type="file"
          accept=".xlsx,.xls,.csv"
          (change)="onFileSelected($event)"
          hidden>

        <div *ngIf="!selectedFile" class="drop-zone-content">
          <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="1.5">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
            <polyline points="17 8 12 3 7 8"/>
            <line x1="12" y1="3" x2="12" y2="15"/>
          </svg>
          <p class="drop-text">Drag & drop your Excel file here</p>
          <p class="drop-subtext">or click to browse &mdash; .xlsx, .xls, .csv supported</p>
        </div>

        <div *ngIf="selectedFile" class="file-info">
          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#107c10" stroke-width="1.5">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
            <polyline points="14 2 14 8 20 8"/>
            <line x1="16" y1="13" x2="8" y2="13"/>
            <line x1="16" y1="17" x2="8" y2="17"/>
            <polyline points="10 9 9 9 8 9"/>
          </svg>
          <div>
            <p class="file-name">{{ selectedFile.name }}</p>
            <p class="file-size">{{ formatFileSize(selectedFile.size) }}</p>
          </div>
          <button class="remove-btn" (click)="removeFile($event)" title="Remove file">&times;</button>
        </div>
      </div>

      <div class="actions">
        <button
          class="assess-btn"
          [disabled]="!selectedFile || isLoading"
          (click)="uploadFile()">
          <span *ngIf="!isLoading">Assess Resources</span>
          <span *ngIf="isLoading" class="loading">
            <span class="spinner"></span> Processing...
          </span>
        </button>
        <button
          *ngIf="selectedFile"
          class="reset-btn"
          (click)="clearAll()">
          Clear
        </button>
      </div>

      <div *ngIf="errorMessage" class="error-bar">
        <span>&#9888;</span> {{ errorMessage }}
      </div>
    </div>
  `,
  styles: [`
    .upload-card {
      background: #fff;
      border-radius: 8px;
      box-shadow: 0 1.6px 3.6px rgba(0,0,0,0.13), 0 0.3px 0.9px rgba(0,0,0,0.1);
      padding: 28px 32px;
      margin-bottom: 24px;
    }
    .upload-header h2 {
      font-size: 20px;
      font-weight: 600;
      color: #323130;
      margin-bottom: 4px;
    }
    .upload-header p {
      font-size: 14px;
      color: #605e5c;
      margin-bottom: 20px;
    }
    .drop-zone {
      border: 2px dashed #c8c6c4;
      border-radius: 6px;
      padding: 40px 20px;
      text-align: center;
      cursor: pointer;
      transition: all 0.2s;
      background: #faf9f8;
    }
    .drop-zone:hover, .drop-zone.drag-over {
      border-color: #0078d4;
      background: #f0f6ff;
    }
    .drop-zone.has-file {
      border-style: solid;
      border-color: #107c10;
      background: #f1faf1;
      padding: 20px;
    }
    .drop-zone-content { display: flex; flex-direction: column; align-items: center; gap: 8px; }
    .drop-text { font-size: 16px; font-weight: 600; color: #323130; }
    .drop-subtext { font-size: 13px; color: #a19f9d; }
    .file-info { display: flex; align-items: center; gap: 12px; }
    .file-name { font-weight: 600; font-size: 14px; color: #323130; }
    .file-size { font-size: 12px; color: #a19f9d; }
    .remove-btn {
      margin-left: auto;
      background: none;
      border: none;
      font-size: 20px;
      color: #a4262c;
      cursor: pointer;
      padding: 4px 8px;
      border-radius: 4px;
    }
    .remove-btn:hover { background: #fde7e9; }
    .actions { display: flex; gap: 12px; margin-top: 20px; }
    .assess-btn {
      background: #0078d4;
      color: #fff;
      border: none;
      padding: 10px 28px;
      border-radius: 4px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: background 0.15s;
    }
    .assess-btn:hover:not(:disabled) { background: #106ebe; }
    .assess-btn:disabled { opacity: 0.5; cursor: not-allowed; }
    .reset-btn {
      background: #fff;
      color: #323130;
      border: 1px solid #8a8886;
      padding: 10px 20px;
      border-radius: 4px;
      font-size: 14px;
      cursor: pointer;
    }
    .reset-btn:hover { background: #f3f2f1; }
    .loading { display: flex; align-items: center; gap: 8px; }
    .spinner {
      width: 16px; height: 16px;
      border: 2px solid rgba(255,255,255,0.3);
      border-top-color: #fff;
      border-radius: 50%;
      animation: spin 0.6s linear infinite;
      display: inline-block;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    .error-bar {
      margin-top: 16px;
      padding: 10px 16px;
      background: #fde7e9;
      color: #a4262c;
      border-radius: 4px;
      font-size: 13px;
      display: flex;
      align-items: center;
      gap: 8px;
    }
  `]
})
export class UploadComponent {
  @Input() mode: 'subscription' | 'region' | 'jio' | 'aws' | 'gcp' = 'subscription';
  @Output() assessmentComplete = new EventEmitter<AssessmentResponse>();
  @Output() fileSelected = new EventEmitter<File>();
  @Output() reset = new EventEmitter<void>();

  selectedFile: File | null = null;
  isDragOver = false;
  isLoading = false;
  errorMessage = '';

  constructor(private migrationService: MigrationService) {}

  onDragOver(event: DragEvent): void {
    event.preventDefault();
    event.stopPropagation();
    this.isDragOver = true;
  }

  onDragLeave(event: DragEvent): void {
    event.preventDefault();
    event.stopPropagation();
    this.isDragOver = false;
  }

  onDrop(event: DragEvent): void {
    event.preventDefault();
    event.stopPropagation();
    this.isDragOver = false;
    const files = event.dataTransfer?.files;
    if (files && files.length > 0) {
      this.selectFile(files[0]);
    }
  }

  onFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      this.selectFile(input.files[0]);
    }
  }

  selectFile(file: File): void {
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (!validExtensions.includes(ext)) {
      this.errorMessage = 'Invalid file type. Please upload an .xlsx, .xls, or .csv file.';
      return;
    }
    if (file.size > 10 * 1024 * 1024) {
      this.errorMessage = 'File is too large. Maximum size is 10 MB.';
      return;
    }
    this.selectedFile = file;
    this.errorMessage = '';
    this.fileSelected.emit(file);
  }

  removeFile(event: Event): void {
    event.stopPropagation();
    this.selectedFile = null;
    this.errorMessage = '';
  }

  uploadFile(): void {
    if (!this.selectedFile) return;
    this.isLoading = true;
    this.errorMessage = '';

    const request$ = this.mode === 'aws'
      ? this.migrationService.assessAwsFile(this.selectedFile)
      : this.mode === 'gcp'
        ? this.migrationService.assessGcpFile(this.selectedFile)
        : this.mode === 'jio'
          ? this.migrationService.assessJioFile(this.selectedFile)
          : this.mode === 'region'
            ? this.migrationService.assessRegionFile(this.selectedFile)
            : this.migrationService.assessFile(this.selectedFile);

    request$.subscribe({
      next: (result) => {
        this.isLoading = false;
        this.assessmentComplete.emit(result);
      },
      error: (err) => {
        this.isLoading = false;
        this.errorMessage = err.error?.error || 'An error occurred while processing the file.';
      }
    });
  }

  clearAll(): void {
    this.selectedFile = null;
    this.errorMessage = '';
    this.reset.emit();
  }

  formatFileSize(bytes: number): string {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  }
}
