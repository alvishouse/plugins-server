import { OpenAPIRegistry } from '@asteasolutions/zod-to-openapi';
import { Readability } from '@mozilla/readability';
import * as cheerio from 'cheerio';
import express, { Request, Response, Router } from 'express';
import got from 'got';
import { StatusCodes } from 'http-status-codes';
import { JSDOM } from 'jsdom';

import { createApiResponse } from '@/api-docs/openAPIResponseBuilders';
import { ResponseStatus, ServiceResponse } from '@/common/models/serviceResponse';
import { handleServiceResponse } from '@/common/utils/httpHandlers';

import { WebPageReaderRequestParamSchema, WebPageReaderResponseSchema } from './webPageReaderModel';

export const articleReaderRegistry = new OpenAPIRegistry();
articleReaderRegistry.register('Web Page Reader', WebPageReaderResponseSchema);

const removeUnwantedElements = (_cheerio: any) => {
  const elementsToRemove = [
    'footer',
    'header',
    'nav',
    'script',
    'style',
    'link',
    'meta',
    'noscript',
    'img',
    'picture',
    'video',
    'audio',
    'iframe',
    'object',
    'embed',
    'param',
    'track',
    'source',
    'canvas',
    'map',
    'area',
    'svg',
    'math',
  ];

  elementsToRemove.forEach((element) => _cheerio(element).remove());
};

const fetchAndCleanContent = async (url: string) => {
  const { body } = await got(url);
  const $ = cheerio.load(body);
  const title = $('title').text();
  removeUnwantedElements($);
  const doc = new JSDOM($.text(), {
    url: url,
  });
  const reader = new Readability(doc.window.document);
  const article = reader.parse();

  return { title, content: article ? article.textContent : '' };
};

export const webPageReaderRouter: Router = (() => {
  const router = express.Router();

  articleReaderRegistry.registerPath({
    method: 'get',
    path: '/web-page-reader/get-content',
    tags: ['Web Page Reader'],
    request: {
      query: WebPageReaderRequestParamSchema,
    },
    responses: createApiResponse(WebPageReaderResponseSchema, 'Success'),
  });

  router.get('/get-content', async (_req: Request, res: Response) => {
    const { url } = _req.query;

    if (typeof url !== 'string') {
      return new ServiceResponse(ResponseStatus.Failed, 'URL must be a string', null, StatusCodes.BAD_REQUEST);
    }

    try {
      const content = await fetchAndCleanContent(url);
      const serviceResponse = new ServiceResponse(
        ResponseStatus.Success,
        'Content fetched successfully',
        content,
        StatusCodes.OK
      );
      return handleServiceResponse(serviceResponse, res);
    } catch (error) {
      console.error(`Error fetching content ${(error as Error).message}`);
      const errorMessage = `Error fetching content ${(error as Error).message}`;
      const serviceResponse = new ServiceResponse(
        ResponseStatus.Failed,
        errorMessage,
        null,
        StatusCodes.INTERNAL_SERVER_ERROR
      );
      return handleServiceResponse(serviceResponse, res);
    }
  });

  return router;
})();
