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
   * Deletes existing sheet if present, injects fresh sheet, attaches onChanged listener for live highlighting.
   */
  async downloadAndInject(
    entityName: string, 
    whereClause: string = '', 
    transformerBean: string = ''
  ) {
    const payload = { 
      entityName, 
      whereClause, 
      customTransformerBean: transformerBean.trim()
    };

    const res = await lastValueFrom(this.http.post<any>(`${this.apiBase}/download`, payload));

    await Excel.run(async (context: any) => {
      const workbook = context.workbook;

      // 1. CLEANUP: Delete old sheet if it already exists
      const existingSheet = workbook.worksheets.getItemOrNullObject(entityName);
      await context.sync();
      if (!existingSheet.isNullObject) {
        existingSheet.delete();
        await context.sync();
      }

      // 2. INJECT: Add new sheet from Base64 binary
      const insertedSheets = workbook.insertWorksheetsFromBase64(res.base64, {
        sheetNamesToInsert: [res.sourceSheetName],
        positionType: Excel.WorksheetPositionType.end
      });
      await context.sync();

      const sheet = workbook.worksheets.getItem(insertedSheets.value[0]);
      sheet.name = entityName;

      // 3. TABLE FORMATTING
      const usedRange = sheet.getUsedRange();
      const table = sheet.tables.add(usedRange, true);
      table.name = `Table_${entityName}`;
      usedRange.format.autofitColumns();

      // 4. LIVE EVENT LISTENER: Highlight edited cells/inserted rows in yellow
      sheet.onChanged.add(async (event: any) => {
        if (event.changeType === 'RangeEdited' || event.changeType === 'RowInserted') {
          await Excel.run(async (ctx: any) => {
            const currentSheet = ctx.workbook.worksheets.getItem(entityName);
            const editedRange = currentSheet.getRange(event.address);
            editedRange.format.fill.color = '#FFF2CC'; // Soft yellow highlight
            await ctx.sync();
          });
        }
      });

      await context.sync();
    });
  }

  /**
   * UPLOAD MODIFIED ROWS
   * Reads grid, sends updates to Spring Boot, and re-downloads fresh sheet (clearing yellow highlights).
   */
  async uploadModifiedRows(entityName: string, transformerBean: string = '') {
    let modifiedRows: any[] = [];

    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getItem(entityName);
      const usedRange = sheet.getUsedRange(true);
      usedRange.load(['values']);
      await context.sync();

      const data: any[][] = usedRange.values;
      if (!data || data.length <= 1) return;

      const headers: string[] = data[0].map((h: any) =>
        h ? String(h).trim().toLowerCase() : ''
      );

      for (let i = 1; i < data.length; i++) {
        const rowData: Record<string, any> = {};
        let hasContent = false;

        headers.forEach((header, colIdx) => {
          if (header) {
            const val = data[i][colIdx];
            if (val !== null && val !== undefined && val !== '') {
              rowData[header] = val;
              hasContent = true;
            } else {
              rowData[header] = null;
            }
          }
        });

        if (hasContent) {
          modifiedRows.push(rowData);
        }
      }
    });

    if (modifiedRows.length === 0) return;

    // Send updates to Java
    const uploadUrl = `${this.apiBase}/upload?entityName=${encodeURIComponent(entityName)}`;
    await lastValueFrom(this.http.post<any>(uploadUrl, modifiedRows));

    // Refresh sheet: Deletes old highlighted sheet and injects fresh clean data from Java
    await this.downloadAndInject(entityName, '', transformerBean);
  }
}