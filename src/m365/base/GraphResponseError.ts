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