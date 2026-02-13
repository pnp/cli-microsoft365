import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015009_FILE_config_sass_json extends FileAddRemoveRule {
  constructor(options: { add: boolean; contents: string }) {
    super({ filePath: './config/sass.json', ...options });
  }

  get id(): string {
    return 'FN015009';
  }
}
