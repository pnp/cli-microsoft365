export interface TsLintJson {
  $schema?: string;
  lintConfig?: {
    rules?: {
      [key: string]: boolean;
    }
  }
}