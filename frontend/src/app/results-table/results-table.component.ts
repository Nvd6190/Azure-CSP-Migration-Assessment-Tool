import { Component, Input, OnChanges } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { AssessmentResponse, AssessmentResource } from '../models/assessment.model';
import { MigrationService } from '../services/migration.service';

@Component({
  selector: 'app-results-table',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
    <!-- Summary Cards (clickable filters) -->
    <div class="summary-row">
      <div class="summary-card total" [class.active]="activeFilter === 'all'" (click)="setFilter('all')">
        <div class="card-value">{{ data.summary.total }}</div>
        <div class="card-label">Total Resources</div>
      </div>
      <div class="summary-card yes" [class.active]="activeFilter === 'Yes'" (click)="setFilter('Yes')">
        <div class="card-value">{{ data.summary.yes }}</div>
        <div class="card-label">{{ mode === 'jio' ? 'Available' : 'Can Move' }}</div>
      </div>
      <div class="summary-card no" [class.active]="activeFilter === 'No'" (click)="setFilter('No')">
        <div class="card-value">{{ data.summary.no }}</div>
        <div class="card-label">{{ mode === 'jio' ? 'Not Available' : 'Cannot Move' }}</div>
      </div>
      <div class="summary-card review" [class.active]="activeFilter === 'Review'" (click)="setFilter('Review')">
        <div class="card-value">{{ data.summary.review }}</div>
        <div class="card-label">Needs Review</div>
      </div>
      <div class="summary-card conditional" *ngIf="mode === 'subscription'" [class.active]="activeFilter === 'Conditional'" (click)="setFilter('Conditional')">
        <div class="card-value">{{ data.summary.conditional || 0 }}</div>
        <div class="card-label">Conditional</div>
      </div>
    </div>

    <!-- Controls -->
    <div class="controls-bar">
      <div class="controls-right">
        <div class="search-box">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#605e5c" stroke-width="2">
            <circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/>
          </svg>
          <input
            type="text"
            placeholder="Search by name, type, or group..."
            [(ngModel)]="searchQuery"
            (ngModelChange)="applyFilters()">
        </div>
        <button class="download-btn" (click)="downloadExcel()">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
            <polyline points="7 10 12 15 17 10"/>
            <line x1="12" y1="15" x2="12" y2="3"/>
          </svg>
          Download Excel
        </button>
      </div>
    </div>

    <!-- Results Table -->
    <div class="table-container">
      <table>
        <thead>
          <tr>
            <th class="col-num">#</th>
            <th class="col-name">Name</th>
            <th class="col-type">Resource Type</th>
            <th class="col-rg">Resource Group</th>
            <th class="col-location">Location</th>
            <th *ngIf="mode === 'jio'" class="col-india">India Region</th>
            <th class="col-status">{{ mode === 'jio' ? 'Jio Availability' : mode === 'region' ? 'Region Move' : 'Subscription Move' }}</th>
            <th class="col-remarks">Remarks</th>
          </tr>
        </thead>
        <tbody>
          <tr *ngFor="let resource of filteredResources; let i = index">
            <td class="col-num">{{ i + 1 }}</td>
            <td class="col-name" [title]="getField(resource, 'NAME')">{{ getField(resource, 'NAME') }}</td>
            <td class="col-type" [title]="getResourceType(resource)">{{ getResourceType(resource) }}</td>
            <td class="col-rg" [title]="getField(resource, 'RESOURCE GROUP')">{{ getField(resource, 'RESOURCE GROUP') }}</td>
            <td class="col-location">{{ getField(resource, 'LOCATION') }}</td>
            <td *ngIf="mode === 'jio'" class="col-india">
              <span class="badge" [ngClass]="resource['INDIA REGION'] === 'Yes' ? 'badge-yes' : resource['INDIA REGION'] === 'No' ? 'badge-no' : 'badge-review'">
                {{ resource['INDIA REGION'] === 'Yes' ? '✓' : resource['INDIA REGION'] === 'No' ? '✗' : '—' }} {{ resource['INDIA REGION'] || '—' }}
              </span>
            </td>
            <td class="col-status">
              <span class="badge" [ngClass]="getBadgeClass(resource)">
                {{ getBadgeIcon(resource) }} {{ getStatusField(resource) }}
              </span>
            </td>
            <td class="col-remarks" [title]="resource['REMARKS'] || ''">{{ resource['REMARKS'] }}</td>
          </tr>
          <tr *ngIf="filteredResources.length === 0">
            <td [attr.colspan]="mode === 'jio' ? 8 : 7" class="empty-row">No resources match your filter.</td>
          </tr>
        </tbody>
      </table>
    </div>
  `,
  styles: [`
    .summary-row {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
      gap: 16px;
      margin-bottom: 20px;
    }
    .summary-card {
      background: #fff;
      border-radius: 8px;
      padding: 20px;
      box-shadow: 0 1.6px 3.6px rgba(0,0,0,0.13), 0 0.3px 0.9px rgba(0,0,0,0.1);
      border-top: 4px solid;
      cursor: pointer;
      transition: transform 0.15s, box-shadow 0.15s;
      user-select: none;
    }
    .summary-card:hover {
      transform: translateY(-2px);
      box-shadow: 0 3.2px 7.2px rgba(0,0,0,0.18), 0 0.6px 1.8px rgba(0,0,0,0.13);
    }
    .summary-card.active {
      outline: 3px solid;
      outline-offset: -3px;
      transform: translateY(-2px);
      box-shadow: 0 3.2px 7.2px rgba(0,0,0,0.18), 0 0.6px 1.8px rgba(0,0,0,0.13);
    }
    .summary-card.total { border-top-color: #0078d4; }
    .summary-card.total.active { outline-color: #0078d4; }
    .summary-card.yes { border-top-color: #107c10; }
    .summary-card.yes.active { outline-color: #107c10; }
    .summary-card.no { border-top-color: #a4262c; }
    .summary-card.no.active { outline-color: #a4262c; }
    .summary-card.review { border-top-color: #ca5010; }
    .summary-card.review.active { outline-color: #ca5010; }
    .summary-card.conditional { border-top-color: #8764b8; }
    .summary-card.conditional.active { outline-color: #8764b8; }
    .card-value { font-size: 32px; font-weight: 700; color: #323130; }
    .card-label { font-size: 13px; color: #605e5c; margin-top: 4px; }

    .controls-bar {
      display: flex;
      justify-content: flex-end;
      align-items: center;
      flex-wrap: wrap;
      gap: 12px;
      margin-bottom: 16px;
    }
    .controls-right { display: flex; gap: 10px; align-items: center; }
    .search-box {
      display: flex;
      align-items: center;
      gap: 8px;
      background: #fff;
      border: 1px solid #c8c6c4;
      border-radius: 4px;
      padding: 6px 12px;
      min-width: 260px;
    }
    .search-box input {
      border: none;
      outline: none;
      font-size: 13px;
      color: #323130;
      width: 100%;
      background: transparent;
    }
    .download-btn {
      display: flex;
      align-items: center;
      gap: 6px;
      background: #107c10;
      color: #fff;
      border: none;
      padding: 8px 18px;
      border-radius: 4px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      white-space: nowrap;
    }
    .download-btn:hover { background: #0b6a0b; }

    .table-container {
      background: #fff;
      border-radius: 8px;
      overflow-x: auto;
      box-shadow: 0 1.6px 3.6px rgba(0,0,0,0.13), 0 0.3px 0.9px rgba(0,0,0,0.1);
    }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    thead th {
      background: #faf9f8;
      text-align: left;
      padding: 10px 12px;
      font-weight: 600;
      color: #323130;
      border-bottom: 1px solid #edebe9;
      white-space: nowrap;
      position: sticky;
      top: 0;
    }
    tbody td {
      padding: 10px 12px;
      border-bottom: 1px solid #f3f2f1;
      color: #323130;
      max-width: 200px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    tbody tr:hover { background: #f3f2f1; }

    .col-num { width: 40px; text-align: center; color: #a19f9d; }
    .col-name { min-width: 160px; }
    .col-type { min-width: 160px; }
    .col-rg { min-width: 140px; }
    .col-location { min-width: 100px; }
    .col-status { min-width: 150px; }
    .col-remarks { min-width: 200px; max-width: 300px; }

    .badge {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 3px 10px;
      border-radius: 12px;
      font-size: 12px;
      font-weight: 600;
      white-space: nowrap;
    }
    .badge-yes { background: #dff6dd; color: #107c10; }
    .badge-no { background: #fde7e9; color: #a4262c; }
    .badge-conditional { background: #e8e0f0; color: #8764b8; }
    .badge-review { background: #fff4ce; color: #ca5010; }

    .empty-row {
      text-align: center;
      color: #a19f9d;
      padding: 40px 12px !important;
      font-size: 14px;
    }

    @media (max-width: 768px) {
      .summary-row { grid-template-columns: repeat(2, 1fr); }
      .controls-bar { flex-direction: column; align-items: stretch; }
      .controls-right { flex-direction: column; }
      .search-box { min-width: unset; }
    }
  `]
})
export class ResultsTableComponent implements OnChanges {
  @Input() data!: AssessmentResponse;
  @Input() mode: 'subscription' | 'region' | 'jio' = 'subscription';

  searchQuery = '';
  activeFilter = 'all';
  filteredResources: AssessmentResource[] = [];
  filterTabs: { label: string; value: string; count: number }[] = [];

  private updateFilterTabs(): void {
    const conditionalCount = this.data.summary.conditional || 0;
    this.filterTabs = [
      { label: 'All', value: 'all', count: this.data.summary.total },
      { label: 'Can Move', value: 'Yes', count: this.data.summary.yes },
      { label: 'Cannot Move', value: 'No', count: this.data.summary.no },
      { label: 'Review', value: 'Review', count: this.data.summary.review },
      { label: 'Conditional', value: 'Conditional', count: conditionalCount }
    ];
  }

  trackTab(_index: number, tab: { value: string }): string {
    return tab.value;
  }

  constructor(private migrationService: MigrationService) {}

  ngOnChanges(): void {
    this.updateFilterTabs();
    this.applyFilters();
  }

  setFilter(value: string): void {
    this.activeFilter = value;
    this.applyFilters();
  }

  applyFilters(): void {
    let resources = this.data.resources;

    if (this.activeFilter !== 'all') {
      resources = resources.filter(r => this.getStatusField(r) === this.activeFilter);
    }

    if (this.searchQuery.trim()) {
      const q = this.searchQuery.toLowerCase();
      resources = resources.filter(r => {
        const name = this.getField(r, 'NAME').toLowerCase();
        const type = this.getResourceType(r).toLowerCase();
        const rg = this.getField(r, 'RESOURCE GROUP').toLowerCase();
        return name.includes(q) || type.includes(q) || rg.includes(q);
      });
    }

    this.filteredResources = resources;
  }

  getField(resource: AssessmentResource, field: string): string {
    const keys = Object.keys(resource);
    // Exact case-insensitive match
    const match = keys.find(k => k.toUpperCase() === field.toUpperCase());
    if (match) return String(resource[match] ?? '');
    // Underscore variant
    const underscored = field.replace(/ /g, '_');
    const matchUnderscore = keys.find(k => k.toUpperCase() === underscored.toUpperCase());
    if (matchUnderscore) return String(resource[matchUnderscore] ?? '');
    // Broader match: look for key that contains the field word
    // e.g., field="TYPE" matches "Resource Type", "RESOURCE TYPE", "Azure Resource Type"
    const fieldUpper = field.toUpperCase();
    const broadMatch = keys.find(k => {
      const kUpper = k.toUpperCase();
      return kUpper.includes(fieldUpper) || fieldUpper.includes(kUpper);
    });
    return broadMatch ? String(resource[broadMatch] ?? '') : '';
  }

  getResourceType(resource: AssessmentResource): string {
    // Priority: NORMALIZED TYPE (ARM format added by assessment) > Resource Type > Type
    const normalized = (resource as any)['NORMALIZED TYPE'];
    if (normalized) return String(normalized);
    const keys = Object.keys(resource);
    // Try "Resource Type" first (ARM format in Azure exports)
    const rtKey = keys.find(k => k.toUpperCase() === 'RESOURCE TYPE');
    if (rtKey) return String(resource[rtKey] ?? '');
    // Fallback to "Type"
    return this.getField(resource, 'TYPE');
  }

  getStatusField(resource: AssessmentResource): string {
    if (this.mode === 'jio') {
      return (resource as any)['JIO REGION AVAILABLE'] || '';
    }
    if (this.mode === 'region') {
      return (resource as any)['REGION MOVE SUPPORTED'] || resource['SUBSCRIPTION MOVE SUPPORTED'] || '';
    }
    return resource['SUBSCRIPTION MOVE SUPPORTED'] || '';
  }

  getBadgeClass(resource: AssessmentResource): string {
    const status = this.getStatusField(resource);
    if (status === 'Yes') return 'badge-yes';
    if (status === 'No') return 'badge-no';
    if (status === 'Conditional') return 'badge-conditional';
    return 'badge-review';
  }

  getBadgeIcon(resource: AssessmentResource): string {
    const status = this.getStatusField(resource);
    if (status === 'Yes') return '\u2705';
    if (status === 'No') return '\u274C';
    if (status === 'Conditional') return '\u2139\uFE0F';
    return '\u26A0\uFE0F';
  }

  downloadExcel(): void {
    const url = this.migrationService.getDownloadUrl(this.data.downloadId);
    window.open(url, '_blank');
  }
}
