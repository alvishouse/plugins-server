import { OpenAPIRegistry } from '@asteasolutions/zod-to-openapi';
import * as ExcelJS from 'exceljs';
import express, { Request, Response, Router } from 'express';
import fs from 'fs';
import { StatusCodes } from 'http-status-codes';
import cron from 'node-cron';
import path from 'path';

import { createApiRequestBody } from '@/api-docs/openAPIRequestBuilders';
import { createApiResponse } from '@/api-docs/openAPIResponseBuilders';
import { ResponseStatus, ServiceResponse } from '@/common/models/serviceResponse';
import { handleServiceResponse } from '@/common/utils/httpHandlers';

import { ExcelGeneratorRequestBodySchema, ExcelGeneratorResponseSchema } from './excelGeneratorModel';

export const COMPRESS = true;
export const excelGeneratorRegistry = new OpenAPIRegistry();

excelGeneratorRegistry.register('ExcelGenerator', ExcelGeneratorResponseSchema);
excelGeneratorRegistry.registerPath({
  method: 'post',
  path: '/excel-generator/generate',
  tags: ['Excel Generator'],
  request: {
    body: createApiRequestBody(ExcelGeneratorRequestBodySchema, 'application/json'),
  },
  responses: createApiResponse(ExcelGeneratorResponseSchema, 'Success'),
});

// Create folder to contains generated files
const exportsDir = path.join(__dirname, '../../..', 'excel-exports');

if (!fs.existsSync(exportsDir)) {
  fs.mkdirSync(exportsDir, { recursive: true });
}

// Cron job to delete files older than 1 hour
cron.schedule('0 * * * *', () => {
  const now = Date.now();
  const oneHour = 60 * 60 * 1000;
  fs.readdir(exportsDir, (err, files) => {
    if (err) return;
    files.forEach((file) => {
      const filePath = path.join(exportsDir, file);
      fs.stat(filePath, (err, stats) => {
        if (err) return;
        if (now - stats.mtime.getTime() > oneHour) {
          fs.unlink(filePath, () => {});
        }
      });
    });
  });
});

const serverUrl = process.env.RENDER_EXTERNAL_URL || 'http://localhost:3000';

function columnLetterToNumber(letter: string): number {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column = column * 26 + letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return column;
}

function autoFitColumns(
  worksheet: ExcelJS.Worksheet,
  startRow: number,
  rows: any[],
  numColumns: number,
  startCol: number
): void {
  for (let colIdx = 0; colIdx < numColumns; colIdx++) {
    let maxLength = 0;
    rows.forEach((row) => {
      const cellValue = row[colIdx];
      if (cellValue != null) {
        maxLength = Math.max(maxLength, String(cellValue).length);
      }
    });
    const headerCell = worksheet.getCell(startRow, startCol + colIdx).value;
    if (headerCell != null) {
      maxLength = Math.max(maxLength, String(headerCell).length);
    }
    worksheet.getColumn(startCol + colIdx).width = maxLength + 2;
  }
}

export function execGenExcelFuncs(sheetsData: any[], excelConfigs: any): string {
  const workbook = new ExcelJS.Workbook();
  const borderConfigs = excelConfigs.borderStyle
    ? {
        top: { style: excelConfigs.borderStyle },
        left: { style: excelConfigs.borderStyle },
        bottom: { style: excelConfigs.borderStyle },
        right: { style: excelConfigs.borderStyle },
      }
    : {};
  const titleAlignment = { horizontal: 'center', vertical: 'middle', wrapText: excelConfigs.wrapText };
  const titleFont = { name: excelConfigs.fontFamily, bold: true, size: excelConfigs.tableTitleFontSize };
  const headerFont = { name: excelConfigs.fontFamily, bold: true, size: excelConfigs.headerFontSize };
  const headerAlign = { wrapText: excelConfigs.wrapText, horizontal: 'center', vertical: 'middle' };
  const cellFont = { name: excelConfigs.fontFamily, size: excelConfigs.fontSize };
  const cellAlign = { wrapText: excelConfigs.wrapText };

  sheetsData.forEach(({ sheetName, tables }) => {
    const worksheet = workbook.addWorksheet(sheetName);
    tables.forEach(({ startCell = 'A1', title, rows = [], columns = [], skipHeader }) => {
      const startCol = columnLetterToNumber(startCell[0]);
      const startRow = parseInt(startCell.slice(1));
      let rowIndex = startRow;

      if (title) {
        const cell = worksheet.getCell(rowIndex, startCol);
        cell.value = title;
        worksheet.mergeCells(rowIndex, startCol, rowIndex, startCol + columns.length - 1);
        cell.alignment = titleAlignment;
        cell.font = titleFont;
        cell.border = borderConfigs;
        rowIndex++;
      }

      if (!skipHeader && columns) {
        columns.forEach((col, colIdx) => {
          const cell = worksheet.getCell(rowIndex, startCol + colIdx);
          cell.value = col.name;
          cell.alignment = headerAlign;
          cell.font = headerFont;
          cell.border = borderConfigs;
        });
        rowIndex++;
      }

      const columnTypes = columns.map((col: any) => col.type) || [];
      const columnFormats =
        columns.map((col: any) => {
          switch (col.type) {
            case 'number': return col.format || undefined;
            case 'percent': return col.format || '0.00%';
            case 'currency': return col.format || '$#,##0';
            case 'date': return col.format || undefined;
            default: return undefined;
          }
        }) || [];

      rows.forEach((rowData) => {
        rowData.forEach((cellData, colIdx) => {
          const { type = 'static_value', value } = cellData;
          const valueType = columnTypes[colIdx];
          const format = columnFormats[colIdx];
          const cell = worksheet.getCell(rowIndex, startCol + colIdx);
          let cellValue: any = value != null ? value : '';

          if (type === 'formula') {
            cell.value = { formula: cellValue };
            if (['percent', 'currency', 'number', 'date'].includes(valueType)) {
              cell.numFmt = format;
            }
          } else {
            switch (valueType) {
              case 'number':
                cell.value = !isNaN(Number(cellValue)) ? Math.round(Number(cellValue)) : cellValue;
                cell.numFmt = format || '0';
                break;
              case 'boolean':
                cell.value = Boolean(cellValue);
                break;
              case 'date':
                const parsed = new Date(cellValue);
                cell.value = !isNaN(parsed.getTime()) ? parsed : cellValue;
                cell.numFmt = format || 'yyyy-mm-dd';
                break;
              case 'percent':
              case 'currency':
                cell.value = !isNaN(Number(cellValue)) ? Number(cellValue) : cellValue;
                cell.numFmt = format;
                break;
              default:
                cell.value = String(cellValue);
            }
          }
          cell.font = cellFont;
          cell.border = borderConfigs;
          cell.alignment = cellAlign;
        });
        rowIndex++;
      });

      if (excelConfigs.autoFilter) {
        const lastCol = startCol + columns.length - 1;
        worksheet.autoFilter = {
          from: { row: startRow + 1, column: startCol },
          to: { row: rowIndex - 1, column: lastCol },
        };
      }

      if (excelConfigs.autoFitColumnWidth) {
        autoFitColumns(worksheet, startRow, rows, columns.length, startCol);
      }
    });
  });

  const fileName = `excel-file-${new Date().toISOString().replace(/\D/gi, '')}.xlsx`;
  const filePath = path.join(exportsDir, fileName);
  workbook.xlsx.writeFile(filePath).catch((err) => console.error('Error writing Excel file', err));
  return fileName;
}

export const excelGeneratorRouter: Router = (() => {
  const router = express.Router();

  router.use('/downloads', express.static(exportsDir));

  router.post('/generate', async (_req: Request, res: Response) => {
    const { sheetsData, excelConfigs } = _req.body;

    if (!sheetsData || !sheetsData.length) {
      const errorRes = new ServiceResponse(
        ResponseStatus.Failed,
        '[Validation Error] Sheets data is required!',
        'Please make sure you have sent the excel sheets content generated from TypingMind.',
        StatusCodes.BAD_REQUEST
      );
      return handleServiceResponse(errorRes, res);
    }

    try {
      const fileName = execGenExcelFuncs(sheetsData, {
        fontFamily: excelConfigs.fontFamily ?? DEFAULT_EXCEL_CONFIGS.fontFamily,
        tableTitleFontSize: excelConfigs.titleFontSize ?? DEFAULT_EXCEL_CONFIGS.tableTitleFontSize,
        headerFontSize: excelConfigs.headerFontSize ?? DEFAULT_EXCEL_CONFIGS.headerFontSize,
        fontSize: excelConfigs.fontSize ?? DEFAULT_EXCEL_CONFIGS.fontSize,
        autoFilter: excelConfigs.autoFilter ?? DEFAULT_EXCEL_CONFIGS.autoFilter,
        borderStyle: excelConfigs.borderStyle || DEFAULT_EXCEL_CONFIGS.borderStyle,
        wrapText: excelConfigs.wrapText ?? DEFAULT_EXCEL_CONFIGS.wrapText,
        autoFitColumnWidth: excelConfigs.autoFitColumnWidth ?? DEFAULT_EXCEL_CONFIGS.autoFitColumnWidth,
      });

      const successRes = new ServiceResponse(
        ResponseStatus.Success,
        'File generated successfully',
        { downloadUrl: `${serverUrl}/excel-generator/downloads/${fileName}` },
        StatusCodes.OK
      );
      return handleServiceResponse(successRes, res);
    } catch (error) {
      const errorMessage = (error as Error).message;
      const errorRes = new ServiceResponse(
        ResponseStatus.Failed,
        `Error ${errorMessage}`,
        'Sorry, we couldn’t generate the Excel file.',
        StatusCodes.INTERNAL_SERVER_ERROR
      );
      return handleServiceResponse(errorRes, res);
    }
  });

  return router;
})();
