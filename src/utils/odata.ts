import { Logger } from "../cli";
import request from "../request";

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

export const odata = {
  async getAllItems<T>(url: string, logger: Logger, metadata?: 'none' | 'minimal' | 'full'): Promise<T[]> {
    let items: T[] = [];

    const requestOptions: any = {
      url: url,
      headers: {
        accept: `application/json;odata.metadata=${metadata ?? 'none'}`
      },
      responseType: 'json'
    };

    const res = await request.get<ODataResponse<T>>(requestOptions);
    items = res.value;

    const nextLink = res['@odata.nextLink'] ?? res.nextLink;
    if (nextLink) {
      const nextPageItems = await odata.getAllItems<T>(nextLink, logger, metadata);
      items = items.concat(nextPageItems);
    }
    
    return items;
  }
};