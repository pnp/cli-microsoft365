export interface AzmgmtResponse<T> {
  nextLink?: string;
  value: T[];
}