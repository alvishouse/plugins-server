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

const exportsDir = path.join(__dirname, '../../..', 'excel-exports');

if (!fs.existsSync(exportsDir)) {
  fs.mkdirSync(exportsDir, { recursive: true });
}

cron.schedule('0 * * * *', () => {
  const now = Date.now();
  const oneHour = 60 * 60 * 1000;
  fs.readdir(exportsDir, (err, files) => {
    if (err) {
      console.error(`Error reading directory ${exportsDir}:`, err);
      return;
    }

    files.forEach((file) => {
      const filePath = path.join(exportsDir, file);
      fs.stat(filePath, (err, stats) => {
        if (err) {
          console.error(`Error getting stats for file ${filePath}:`, err);
          return;
        }

        if (now - stats.mtime.getTime() > oneHour) {
          fs.unlink(filePath, (err) => {
            if (err) {
              console.error(`Error deleting file ${filePath}:`, err);
            } else {
              console.log(`Deleted file: ${filePath}`);
            }
          });
        }
      });
    });
  });
});

const serverUrl = process.env.RENDER_EXTERNAL_URL || 'http://localhost:3000';

interface SheetData {
  sheetName: string;
  tables: {
    title: string;
    startCell: string;
    rows: {
      type: string;
      value: string;
    }[][];
    columns: { name: string; type: string; format: string }[];
    skipHeader: boolean;
  }[];
}

interface ExcelConfig {
  fontFamily: string;
  tableTitleFontSize: number;
  headerFontSize: number;
  fontSize: number;
  autoFitColumnWidth: boolean;
  autoFilter: boolean;
  borderStyle: ExcelJS.BorderStyle | null;
  wrapText: boolean;
}

const DEFAULT_EXCEL_CONFIGS: ExcelConfig = {
  fontFamily: 'Calibri',
  tableTitleFontSize: 13,
  headerFontSize: 11,
  fontSize: 11,
  autoFitColumnWidth: true,
  autoFilter: false,
  wrap
