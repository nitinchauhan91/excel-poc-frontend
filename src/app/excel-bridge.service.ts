import { Injectable, inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { lastValueFrom } from 'rxjs';

declare const Excel: any;

@Injectable({ providedIn: 'root' })
export class ExcelBridgeService {
  private http = inject(HttpClient);
  private apiBase = 'http://localhost:8080/api/generic';

  /**
   * DOWNLOAD & INJECT
   * Fetches Base64 payload from Spring Boot and injects sheet into host workbook natively
   */
  public async downloadAndInject(
    entityName: string, 
    whereClause: string = '', 
    transformerBean: string = ''
  ): Promise<void> {
    const payload = { 
      entityName, 
      whereClause, 
      customTransformerBean: transformerBean.trim()
    };

    const res = await lastValueFrom(this.http.post<any>(`${this.apiBase}/download`, payload));

    await Excel.run(async (context: any) => {
      const workbook = context.workbook;

      // 1. Delete existing worksheet if already loaded to avoid duplicate key errors
      const existingSheet = workbook.worksheets.getItemOrNullObject(entityName);
      await context.sync();
      if (!existingSheet.isNullObject) {
        existingSheet.delete();
        await context.sync();
      }

      // 2. Native C++ Base64 Insertion
      const insertedSheets = workbook.insertWorksheetsFromBase64(res.base64, {
        sheetNamesToInsert: [res.sourceSheetName],
        positionType: Excel.WorksheetPositionType.end
      });
      await context.sync();

      // Get exact sheet reference via item index
      const sheet = workbook.worksheets.getItem(insertedSheets.value[0]);
      sheet.name = entityName;

      // 3. Format dataset into a structured Excel Table
      const usedRange = sheet.getUsedRange();
      const table = sheet.tables.add(usedRange, true);
      table.name = `Table_${entityName}`;
      usedRange.format.autofitColumns();

      // 4. Register live change listener for cell highlighting
      sheet.onChanged.add(async (event: any) => {
        if (event.changeType === 'RangeEdited' || event.changeType === 'RowInserted') {
          await Excel.run(async (ctx: any) => {
            const currentSheet = ctx.workbook.worksheets.getItem(entityName);
            const editedRange = currentSheet.getRange(event.address);
            editedRange.format.fill.color = '#FFF2CC'; // Highlight edits yellow
            await ctx.sync();
          });
        }
      });

      await context.sync();
    });
  }

  /**
   * UPLOAD MODIFIED ROWS (STANDARD BATCH FOR SMALL DATASETS ONLY)
   */
  public async uploadModifiedRows(entityName: string, transformerBean: string = ''): Promise<void> {
    // For 600k datasets, route directly to the chunked streaming method
    return this.streamUploadAndTransform(entityName, transformerBean, 5000);
  }

  /**
   * STREAMING UPLOAD (OPTIMIZED FOR 600K BENCHMARK)
   * Streams micro-batches with matrix-range updates and live UI progress color tracking
   */
  public async streamUploadAndTransform(
    entityName: string, 
    transformerBean: string = '', 
    chunkSize: number = 5000 // 5,000 row chunks optimal for high volume
  ): Promise<void> {
    
    let totalRows = 0;
    let headers: string[] = [];

    // Step 1: Lightweight inspection to fetch total rows and headers
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getItem(entityName);
      const usedRange = sheet.getUsedRange(true);
      usedRange.load(['rowCount', 'columnCount']);
      
      const headerRange = sheet.getRange("1:1").getUsedRange();
      headerRange.load('values');
      
      await context.sync();

      totalRows = usedRange.rowCount;
      if (headerRange.values && headerRange.values[0]) {
        headers = headerRange.values[0].map((h: any) => h ? String(h).trim().toLowerCase() : '');
      }
    });

    if (totalRows <= 1 || headers.length === 0) return;

    const totalPriceColIdx = headers.indexOf('total_price');
    const uploadUrl = `${this.apiBase}/upload-stream?entityName=${encodeURIComponent(entityName)}&transformerBean=${encodeURIComponent(transformerBean)}`;

    // Step 2: Read, Upload, and Highlight in 5,000-row CHUNKS
    for (let startRow = 1; startRow < totalRows; startRow += chunkSize) {
      const currentChunkSize = Math.min(chunkSize, totalRows - startRow);
      let batchRows: any[] = [];

      // A. Load ONLY current 5,000-row range into JS memory & Highlight Yellow
      await Excel.run(async (context: any) => {
        const sheet = context.workbook.worksheets.getItem(entityName);
        
        // Target row index window: startRow to startRow + currentChunkSize
        const chunkRange = sheet.getRangeByIndexes(startRow, 0, currentChunkSize, headers.length);
        chunkRange.load('values');
        
        // Visual indicator: Active processing chunk turns yellow
        chunkRange.format.fill.color = '#FFF2CC'; 
        
        await context.sync();

        // Convert batch range array to JSON payload
        for (let r = 0; r < chunkRange.values.length; r++) {
          const rowValues = chunkRange.values[r];
          const rowData: Record<string, any> = {};
          let hasData = false;

          headers.forEach((header, colIdx) => {
            if (header) {
              const val = rowValues[colIdx];
              if (val !== null && val !== undefined && val !== '') {
                rowData[header] = val;
                hasData = true;
              }
            }
          });

          if (hasData) {
            batchRows.push(rowData);
          }
        }
      });

      if (batchRows.length === 0) continue;

      // B. Send micro-batch to Spring Boot
      const processedBatch = await lastValueFrom(
        this.http.post<any[]>(uploadUrl, batchRows)
      );

      // C. Update backend computations (total_price) and turn batch Green
      await Excel.run(async (context: any) => {
        const sheet = context.workbook.worksheets.getItem(entityName);
        const chunkRange = sheet.getRangeByIndexes(startRow, 0, currentChunkSize, headers.length);

        // Matrix update calculated total_price
        if (totalPriceColIdx !== -1 && processedBatch && processedBatch.length > 0) {
          const updateMatrix = processedBatch.map(row => [row['total_price'] ?? null]);
          const targetColumnRange = sheet.getRangeByIndexes(startRow, totalPriceColIdx, currentChunkSize, 1);
          targetColumnRange.values = updateMatrix;
        }

        // Visual indicator: Completed chunk turns green
        chunkRange.format.fill.color = '#D9EAD3';
        
        await context.sync();
      });
    }
  }
}