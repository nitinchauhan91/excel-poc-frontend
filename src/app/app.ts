import { Component, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ExcelBridgeService } from './excel-bridge.service';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './app.html',
  styleUrl: './app.css'
})
export class AppComponent {
  private excelBridgeService = inject(ExcelBridgeService);

  // Form Fields
  entityName = 'product';
  whereClause = '';
  transformerBean = 'productTransformer'; // Can be changed per entity or left blank

  // UI State
  loading = signal<boolean>(false);
  statusMessage = signal<string>('Ready');

  async onDownload() {
    this.loading.set(true);
    this.statusMessage.set('Fetching data from Java backend...');

    try {
      await this.excelBridgeService.downloadAndInject(
        this.entityName,
        this.whereClause,
        this.transformerBean // Passes string value or "" directly
      );
      this.statusMessage.set('Sheet injected successfully!');
    } catch (err: unknown) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      this.statusMessage.set(`Download failed: ${errorMsg}`);
    } finally {
      this.loading.set(false);
    }
  }

  async onUpload() {
    this.loading.set(true);
    this.statusMessage.set('Uploading modified rows to database...');

    try {
      await this.excelBridgeService.uploadModifiedRows(
        this.entityName,
        this.transformerBean
      );
      this.statusMessage.set('Upload complete!');
    } catch (err: unknown) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      this.statusMessage.set(`Upload failed: ${errorMsg}`);
    } finally {
      this.loading.set(false);
    }
  }
}