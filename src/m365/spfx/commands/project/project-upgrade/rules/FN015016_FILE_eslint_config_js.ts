import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015016_FILE_eslint_config_js extends FileAddRemoveRule {
  constructor(options: { add: boolean; contents: string }) {
    super({ filePath: './eslint.config.js', ...options });
  }

  get id(): string {
    return 'FN015016';
  }
}
