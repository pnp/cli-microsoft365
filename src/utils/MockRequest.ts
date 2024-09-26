// Matches Dev Proxy mock request format
export interface MockRequest {
  request: {
    url: string;
    method?: 'GET' | 'POST' | 'PUT' | 'PATCH' | 'DELETE';
    bodyFragment?: string;
  };
  response: {
    statusCode?: number;
    headers?: {
      name: string;
      value: string;
    }[];
    body?: object | string;
  };
}

export interface MockRequests extends Record<string, MockRequest> { }
