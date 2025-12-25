import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015015_FILE_config_typescript_json extends FileAddRemoveRule {
  constructor(options: { add: boolean; contents?: string }) {
    super({
      filePath: './config/typescript.json',
      ...options
    });
  }

  get id(): string {
    return 'FN015015';
  }
}
