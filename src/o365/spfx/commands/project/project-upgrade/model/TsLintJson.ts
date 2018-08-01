export interface TsLintJson {
  $schema: string;
  lintConfig?: {
    rules: {
    "prefer-const"?: boolean;
    "class-name"?: boolean;
    }
  };
}