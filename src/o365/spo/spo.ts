export interface ContextInfo {
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