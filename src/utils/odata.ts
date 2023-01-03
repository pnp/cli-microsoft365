import request, { CliRequestOptions } from "../request";

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
  async getAllItems<T>(url: string, metadata?: 'none' | 'minimal' | 'full'): Promise<T[]> {
    let items: T[] = [];

    const requestOptions: CliRequestOptions = {
      url: url,
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
      const nextPageItems = await odata.getAllItems<T>(nextLink, metadata);
      items = items.concat(nextPageItems);
    }

    return items;
  }
};