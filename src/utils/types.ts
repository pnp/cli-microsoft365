export interface Dictionary<T> {
  [key: string]: T;
}

export interface Hash {
  [key: string]: string;
}

// #region Graph types
export interface GraphBatchRequest {
  requests: GraphBatchRequestItem[];
}

export interface GraphBatchRequestItem {
  id: number | string;
  method: "GET" | "POST" | "PUT" | "PATCH" | "DELETE";
  url: string;
  headers?: { [key: string]: string };
  body?: any;
}

export interface GraphBatchRequestResponse {
  responses: {
    id: number | string;
    status: number;
    headers?: { [key: string]: string };
    body?: any;
  }[];
}
// #endregion