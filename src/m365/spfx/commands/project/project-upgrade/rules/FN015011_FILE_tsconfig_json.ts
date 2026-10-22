import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015011_FILE_tsconfig_json extends FileAddRemoveRule {
  constructor(options: { add: boolean; contents?: string }) {
    super({ filePath: './tsconfig.json', ...options });
  }

  get id(): string {
    return 'FN015011';
  }
}
