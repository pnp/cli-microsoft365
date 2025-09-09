import request, { CliRequestOptions } from "../request.js";

export interface ODataResponse<T> {
  '@odata.nextLink'?: string;
  nextLink?: string;
  value: T[];
}

export interface GraphResponseError {
  error: {
    code: string;
    message: string;
    innerError: {
      "request-id": string;
      date: string;
    }
  }
}

/* eslint-disable no-redeclare */
function getAllItems<T>(url: string): Promise<T[]>;
function getAllItems<T>(options: CliRequestOptions): Promise<T[]>;
function getAllItems<T>(url: string, metadata: 'none' | 'minimal' | 'full'): Promise<T[]>;
/* eslint-enable no-redeclare */

// eslint-disable-next-line no-redeclare
async function getAllItems<T>(param1: unknown, metadata?: 'none' | 'minimal' | 'full'): Promise<T[]> {
  let items: T[] = [];

  const requestOptions: CliRequestOptions = typeof param1 !== 'string' ? param1 as CliRequestOptions : {
    url: param1 as string,
    headers: {
      accept: `application/json;odata.metadata=${metadata ?? 'none'}`,
      'odata-version': '4.0'
    },
    responseType: 'json'
  };

  const res = await request.get<ODataResponse<T>>(requestOptions);
  items = res.value;

  const nextLink = res['@odata.nextLink'] ?? res.nextLink;
  if (nextLink) {
    const nextPageItems = await odata.getAllItems<T>({ ...requestOptions, url: nextLink });
    items = items.concat(nextPageItems);
  }

  return items;
}

export const odata = {
  getAllItems
};