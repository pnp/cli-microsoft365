import { JsonFile } from ".";

export interface TsLintJson extends JsonFile {
  $schema?: string;
  extends?: string;
  lintConfig?: {
    rules?: {
      [key: string]: boolean;
    }
  },
  rulesDirectory?: string[];
}