export interface GraphResponse<T> {
  '@odata.nextLink'?: string;
  value: T[];
}