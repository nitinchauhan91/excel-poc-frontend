import { Component, OnInit, signal, inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
// Fix 1: Correct path to service in the same app/ directory
import { ExcelBridgeService, TimingMetrics } from './excel-bridge.service';

declare const Office: any;

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './app.html'
})
export class AppComponent implements OnInit {
  private excelService = inject(ExcelBridgeService);

  isLoggedIn = signal<boolean>(false);
  loading = signal<boolean>(false);
  statusMessage = signal<string>('');
  
  timings = signal<TimingMetrics | null>(null);

  username = 'admin';
  password = 'password';

  ngOnInit() {
    Office.onReady((info: any) => {
      if (info.host === Office.HostType.Excel) {
        console.log('Office.js host ready inside Excel');
      }
    });
  }

  onLogin() {
    this.isLoggedIn.set(true);
    this.statusMessage.set('Logged in successfully');
  }

  async loadExcelSheet() {
    this.loading.set(true);
    this.statusMessage.set('Fetching and generating report...');
    this.timings.set(null);

    try {
      const metrics: TimingMetrics = await this.excelService.fetchAndInject(
        'http://localhost:8080/api/reports/sales',
        { rowCount: 600000 },
        'Sales_Report'
      );
      
      this.timings.set(metrics);
      this.statusMessage.set('Report loaded successfully!');
    } catch (err: unknown) {
      console.error(err);
      // Fix 2: Safe type casting for unknown error object
      const errorMessage = err instanceof Error ? err.message : String(err);
      this.statusMessage.set('Error loading sheet: ' + errorMessage);
    } finally {
      this.loading.set(false);
    }
  }
}