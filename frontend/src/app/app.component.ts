import { Component, OnInit } from '@angular/core';
import { UploadComponent } from './upload/upload.component';
import { ResultsTableComponent } from './results-table/results-table.component';
import { LandingComponent, AssessmentMode } from './landing/landing.component';
import { AssessmentResponse } from './models/assessment.model';
import { CommonModule } from '@angular/common';
import { MigrationService } from './services/migration.service';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, UploadComponent, ResultsTableComponent, LandingComponent],
  template: `
    <!-- Azure-style topbar -->
    <header class="topbar">
      <div class="topbar-left">
        <img class="azure-logo" src="assets/azure-logo.png" alt="Azure" width="32" height="32"/>
        <div class="topbar-title-group">
          <span class="topbar-title">Azure CSP Migration Assessment Tool</span>
        </div>
        <span class="mode-badge" *ngIf="selectedMode">
          {{ selectedMode === 'jio' ? 'Jio Region' : selectedMode === 'region' ? 'Region Move' : 'Subscription Move' }}
        </span>
      </div>
      <div class="topbar-right">
        <span class="rules-info" *ngIf="rulesInfo">
          <span class="rules-badge" [class.dynamic]="rulesInfo.source && rulesInfo.source !== 'static'">
            {{ rulesInfo.totalRules }} rules
          </span>
          <span class="rules-source">{{ rulesInfo.source && rulesInfo.source !== 'static' ? 'Live' : 'Static' }}</span>
        </span>
        <button class="refresh-btn" (click)="refreshRules()" [disabled]="refreshing" title="Refresh rules from Microsoft">
          <span [class.spinning]="refreshing">&#x21bb;</span>
        </button>
        <button class="jio-update-btn" (click)="jioFileInput.click()" [disabled]="jioUpdating" title="Update Jio availability data">
          <span *ngIf="!jioUpdating">&#x2191; Jio Data</span>
          <span *ngIf="jioUpdating" class="spinning">&#x21bb;</span>
        </button>
        <input #jioFileInput type="file" accept=".xlsx,.xls" (change)="onJioFileSelected($event)" hidden>
        <button *ngIf="selectedMode" class="back-btn" (click)="goHome()" title="Back to home">
          &#8962;
        </button>
      </div>
    </header>

    <main class="main-content">
      <!-- Landing Page: Choose mode -->
      <app-landing
        *ngIf="!selectedMode"
        (modeSelected)="onModeSelected($event)">
      </app-landing>

      <!-- Assessment flow -->
      <ng-container *ngIf="selectedMode">
        <app-upload
          [mode]="selectedMode"
          (assessmentComplete)="onAssessmentComplete($event)"
          (reset)="onReset()">
        </app-upload>

        <app-results-table
          *ngIf="assessmentResult"
          [data]="assessmentResult"
          [mode]="selectedMode">
        </app-results-table>
      </ng-container>
    </main>
  `,
  styles: [`
    .topbar {
      background: #0078d4;
      color: #fff;
      height: 48px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0 20px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.14);
      position: sticky;
      top: 0;
      z-index: 100;
    }
    .topbar-left {
      display: flex;
      align-items: center;
      gap: 12px;
    }
    .topbar-right {
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .azure-logo { flex-shrink: 0; }
    .topbar-title {
      font-size: 16px;
      font-weight: 600;
      letter-spacing: 0.3px;
    }
    .mode-badge {
      background: rgba(255,255,255,0.2);
      padding: 3px 10px;
      border-radius: 12px;
      font-size: 12px;
      font-weight: 600;
    }
    .rules-info {
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 12px;
    }
    .rules-badge {
      background: rgba(255,255,255,0.2);
      padding: 2px 8px;
      border-radius: 10px;
      font-weight: 600;
    }
    .rules-badge.dynamic {
      background: rgba(76, 175, 80, 0.4);
    }
    .rules-source { opacity: 0.85; }
    .refresh-btn, .back-btn {
      background: rgba(255,255,255,0.15);
      border: 1px solid rgba(255,255,255,0.3);
      color: #fff;
      width: 30px;
      height: 30px;
      border-radius: 50%;
      cursor: pointer;
      font-size: 16px;
      display: flex;
      align-items: center;
      justify-content: center;
      transition: background 0.2s;
    }
    .refresh-btn:hover:not(:disabled), .back-btn:hover {
      background: rgba(255,255,255,0.25);
    }
    .refresh-btn:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }
    .jio-update-btn {
      background: rgba(255,255,255,0.15);
      border: 1px solid rgba(255,255,255,0.3);
      color: #fff;
      padding: 4px 12px;
      border-radius: 14px;
      cursor: pointer;
      font-size: 11px;
      font-weight: 600;
      transition: background 0.2s;
      white-space: nowrap;
    }
    .jio-update-btn:hover:not(:disabled) {
      background: rgba(255,255,255,0.25);
    }
    .jio-update-btn:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }
    .spinning {
      display: inline-block;
      animation: spin 1s linear infinite;
    }
    @keyframes spin {
      from { transform: rotate(0deg); }
      to { transform: rotate(360deg); }
    }
    .main-content {
      max-width: 1280px;
      margin: 0 auto;
      padding: 24px 20px;
    }
  `]
})
export class AppComponent implements OnInit {
  selectedMode: AssessmentMode | null = null;
  assessmentResult: AssessmentResponse | null = null;
  rulesInfo: any = null;
  refreshing = false;
  jioUpdating = false;

  constructor(private migrationService: MigrationService) {}

  ngOnInit(): void {
    this.loadRulesInfo();
  }

  loadRulesInfo(): void {
    this.migrationService.getRules().subscribe({
      next: (data) => {
        this.rulesInfo = data;
      }
    });
  }

  refreshRules(): void {
    this.refreshing = true;
    this.migrationService.refreshRules().subscribe({
      next: (data) => {
        this.rulesInfo = { ...this.rulesInfo, ...data };
        this.refreshing = false;
      },
      error: () => {
        this.refreshing = false;
      }
    });
  }

  onModeSelected(mode: AssessmentMode): void {
    this.selectedMode = mode;
  }

  onAssessmentComplete(result: AssessmentResponse): void {
    this.assessmentResult = result;
  }

  onReset(): void {
    this.assessmentResult = null;
  }

  goHome(): void {
    this.selectedMode = null;
    this.assessmentResult = null;
  }

  onJioFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) return;
    const file = input.files[0];
    this.jioUpdating = true;
    this.migrationService.refreshJioData(file).subscribe({
      next: (result) => {
        this.jioUpdating = false;
        this.loadRulesInfo();
        alert(`Jio data updated: ${result.totalServices} services, ${result.totalVMs} VMs`);
      },
      error: (err) => {
        this.jioUpdating = false;
        alert('Failed to update Jio data: ' + (err.error?.error || err.message));
      }
    });
    input.value = '';
  }
}
