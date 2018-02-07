export interface ContextInfo {
  FormDigestTimeoutSeconds: number;
  FormDigestValue: string;
}

export interface ClientSvcResponse extends Array<any | ClientSvcResponseContents> {
}

export interface ClientSvcResponseContents {
  SchemaVersion: string;
  LibraryVersion: string;
  ErrorInfo?: {
    ErrorMessage: string;
    ErrorValue?: string;
    TraceCorrelationId: string;
    ErrorCode: number;
    ErrorTypeName?: string;
  };
  TraceCorrelationId: string;
}

export interface SearchResponse {
  PrimaryQueryResult: {
    RelevantResults: {
      RowCount: number;
      Table: {
        Rows: {
          Cells: {
            Key: string;
            Value: string;
            ValueType: string;
          }[];
        }[];
      };
    }
  }
}