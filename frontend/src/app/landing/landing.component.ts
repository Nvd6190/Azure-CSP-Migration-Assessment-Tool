import { Component, EventEmitter, Output } from '@angular/core';
import { CommonModule } from '@angular/common';

export type AssessmentMode = 'subscription' | 'region' | 'jio';

@Component({
  selector: 'app-landing',
  standalone: true,
  imports: [CommonModule],
  template: `
    <div class="landing-container">
      <div class="landing-header">
        <div class="landing-logo">
          <img src="assets/azure-logo.png" alt="Microsoft Azure" width="72" height="72"/>
        </div>
        <h1>Azure CSP Migration Assessment Tool</h1>
        <p>Choose the type of migration assessment to perform on your Azure resources</p>
      </div>

      <div class="cards-row">
        <!-- Subscription Move Card -->
        <div class="choice-card" (click)="select('subscription')">
          <div class="card-icon subscription-icon">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="1.5">
              <rect x="2" y="3" width="20" height="18" rx="2"/>
              <path d="M8 12h8m-4-4v8"/>
            </svg>
          </div>
          <h2>Subscription Move</h2>
          <p class="card-desc">Assess whether your Azure resources can be moved to a different subscription (e.g., CSP migration)</p>
          <div class="card-details">
            <span class="detail-item">&#10003; Cross-subscription move support</span>
            <span class="detail-item">&#10003; 760+ resource types checked</span>
            <span class="detail-item">&#10003; Official Microsoft data</span>
          </div>
          <button class="card-btn subscription-btn">Start Assessment &rarr;</button>
        </div>

        <!-- Region Move Card -->
        <div class="choice-card" (click)="select('region')">
          <div class="card-icon region-icon">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="1.5">
              <circle cx="12" cy="12" r="10"/>
              <path d="M2 12h20M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/>
            </svg>
          </div>
          <h2>Region Move</h2>
          <p class="card-desc">Assess whether your Azure resources can be moved to a different Azure region (e.g., disaster recovery, compliance)</p>
          <div class="card-details">
            <span class="detail-item">&#10003; Cross-region move support</span>
            <span class="detail-item">&#10003; 750+ resource types checked</span>
            <span class="detail-item">&#10003; Official Microsoft data</span>
          </div>
          <button class="card-btn region-btn">Start Assessment &rarr;</button>
        </div>

        <!-- Jio Region Card -->
        <div class="choice-card" (click)="select('jio')">
          <div class="card-icon jio-icon">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#e60012" stroke-width="1.5">
              <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2z"/>
              <path d="M12 6v6l4 2"/>
              <path d="M8 14h8"/>
            </svg>
          </div>
          <h2>Jio Region Availability</h2>
          <p class="card-desc">Check whether your Azure services are available in the Jio India West (JIO) region</p>
          <div class="card-details">
            <span class="detail-item">&#10003; 220+ services checked</span>
            <span class="detail-item">&#10003; 124 VM series verified</span>
            <span class="detail-item">&#10003; Jio India West region data</span>
          </div>
          <button class="card-btn jio-btn">Start Assessment &rarr;</button>
        </div>
      </div>
    </div>
  `,
  styles: [`
    .landing-container {
      display: flex;
      flex-direction: column;
      align-items: center;
      padding: 40px 20px;
    }
    .landing-header {
      text-align: center;
      margin-bottom: 40px;
    }
    .landing-logo {
      margin-bottom: 16px;
      display: flex;
      justify-content: center;
    }
    .landing-header h1 {
      font-size: 28px;
      font-weight: 700;
      color: #323130;
      margin-bottom: 8px;
    }
    .landing-header p {
      font-size: 16px;
      color: #605e5c;
    }
    .cards-row {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 24px;
      max-width: 1000px;
      width: 100%;
    }
    .choice-card {
      background: #fff;
      border-radius: 12px;
      padding: 32px 28px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      border: 2px solid transparent;
      cursor: pointer;
      transition: all 0.2s ease;
      display: flex;
      flex-direction: column;
      align-items: center;
      text-align: center;
    }
    .choice-card:hover {
      border-color: #0078d4;
      transform: translateY(-4px);
      box-shadow: 0 8px 24px rgba(0,120,212,0.15);
    }
    .card-icon {
      margin-bottom: 20px;
      width: 72px;
      height: 72px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .subscription-icon { background: #e8f4fd; }
    .region-icon { background: #e8f8e8; }
    .jio-icon { background: #fde8e8; }
    .choice-card h2 {
      font-size: 20px;
      font-weight: 700;
      color: #323130;
      margin-bottom: 10px;
    }
    .card-desc {
      font-size: 14px;
      color: #605e5c;
      line-height: 1.5;
      margin-bottom: 20px;
    }
    .card-details {
      display: flex;
      flex-direction: column;
      gap: 6px;
      margin-bottom: 24px;
      align-self: stretch;
    }
    .detail-item {
      font-size: 13px;
      color: #323130;
      text-align: left;
      padding-left: 8px;
    }
    .card-btn {
      padding: 10px 24px;
      border-radius: 6px;
      border: none;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: background 0.15s;
      color: #fff;
    }
    .subscription-btn { background: #0078d4; }
    .subscription-btn:hover { background: #106ebe; }
    .region-btn { background: #107c10; }
    .region-btn:hover { background: #0b6a0b; }
    .jio-btn { background: #e60012; }
    .jio-btn:hover { background: #c4000f; }

    @media (max-width: 700px) {
      .cards-row { grid-template-columns: 1fr; }
    }
  `]
})
export class LandingComponent {
  @Output() modeSelected = new EventEmitter<AssessmentMode>();

  select(mode: AssessmentMode): void {
    this.modeSelected.emit(mode);
  }
}
