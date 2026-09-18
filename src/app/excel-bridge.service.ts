import { Injectable, inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { lastValueFrom } from 'rxjs';

declare const Office: any;
declare const Excel: any;

export interface TimingMetrics {
  javaNetworkMs: number;
  rawImportMs: number;
  renameSheetMs: number;
  formulaEvaluationMs: number;
  totalMs: number;
}

export interface BackendExcelResponse {
  base64: string;
  sourceSheetName: string;
}

@Injectable({
  providedIn: 'root'
})
export class ExcelBridgeService {
  private http = inject(HttpClient);

  public async fetchAndInject(
    apiEndpoint: string,
    requestPayload: any,
    targetSheetName: string
  ): Promise<TimingMetrics> {
    const t0 = performance.now();

    const response = await lastValueFrom(
      this.http.post<BackendExcelResponse>(apiEndpoint, requestPayload)
    );

    const t1 = performance.now();

    if (!response || !response.base64 || !response.sourceSheetName) {
      throw new Error('Invalid backend response structure.');
    }

    let t2 = 0;
    let t3 = 0;
    let t4 = 0;

    await Excel.run(async (context: any) => {
      const workbook = context.workbook;

      context.application.calculationMode = Excel.CalculationMode.manual;
      context.application.suspendScreenUpdatingUntilNextSync();

      const insertedSheetIds = workbook.insertWorksheetsFromBase64(
        response.base64,
        {
          sheetNamesToInsert: [response.sourceSheetName],
          positionType: Excel.WorksheetPositionType.end
        }
      );

      await context.sync();
      t2 = performance.now();

      const sheetIds = insertedSheetIds.value;
      if (!sheetIds || !Array.isArray(sheetIds) || sheetIds.length === 0) {
        throw new Error('No worksheet ID returned after Base64 injection.');
      }

      const insertedSheet = workbook.worksheets.getItem(sheetIds[0]);
      insertedSheet.name = targetSheetName;

      await context.sync();
      t3 = performance.now();

      context.application.calculationMode = Excel.CalculationMode.automatic;
      context.application.calculate(Excel.CalculationType.full);

      await context.sync();
      t4 = performance.now();
    });

    return {
      javaNetworkMs: Math.round(t1 - t0),
      rawImportMs: Math.round(t2 - t1),
      renameSheetMs: Math.round(t3 - t2),
      formulaEvaluationMs: Math.round(t4 - t3),
      totalMs: Math.round(t4 - t0)
    };
  }
}