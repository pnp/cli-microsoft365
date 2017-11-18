export interface RestResponse<T>
{
  d: {
    results : T[];
  }
}