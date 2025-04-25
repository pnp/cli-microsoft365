import type { INetworkModule, NetworkRequestOptions, NetworkResponse } from '@azure/msal-node';
import type { AxiosError, AxiosResponse } from 'axios';
import request, { CliRequestOptions } from '../request.js';

export class MsalNetworkClient implements INetworkModule {
  sendGetRequestAsync<T>(url: string, options?: NetworkRequestOptions): Promise<NetworkResponse<T>> {
    return this.sendRequestAsync(url, 'GET', options);
  }
  sendPostRequestAsync<T>(url: string, options?: NetworkRequestOptions): Promise<NetworkResponse<T>> {
    return this.sendRequestAsync(url, 'POST', options);
  }

  private async sendRequestAsync<T>(
    url: string,
    method: 'GET' | 'POST',
    options: NetworkRequestOptions = {},
  ): Promise<NetworkResponse<T>> {
    const requestOptions: CliRequestOptions = {
      url: url,
      method: method,
      headers: {
        'x-anonymous': true,
        ...options.headers
      },
      data: options.body,
      fullResponse: true
    };

    try {
      const res: AxiosResponse = await request.execute<AxiosResponse>(requestOptions);
      const headersObj: Record<string, string> = {};
      for (const [key, value] of Object.entries(res.headers)) {
        headersObj[key] = typeof value === 'string' ? value : String(value);
      }

      return {
        headers: headersObj,
        body: JSON.parse(res.data),
        status: res.status
      };
    }
    catch (ex) {
      const error = ex as AxiosError;
      const headersObj: Record<string, string> = {};
      if (error.response?.headers) {
        for (const [key, value] of Object.entries(error.response.headers)) {
          headersObj[key] = typeof value === 'string' ? value : String(value);
        }
      }

      return {
        headers: headersObj,
        body: JSON.parse(error.response?.data as string ?? '{}') as any,
        status: error.response?.status ?? 400
      };
    }
  }
}