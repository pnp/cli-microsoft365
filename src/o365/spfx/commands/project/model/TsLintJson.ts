export interface TsLintJson {
  $schema?: string;
  extends?: string;
  lintConfig?: {
    rules?: {
      [key: string]: boolean;
    }
  },
  rulesDirectory?: string[];
}