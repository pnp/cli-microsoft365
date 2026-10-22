import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015003_FILE_tslint_json extends FileAddRemoveRule {
  constructor(options: { add: boolean; contents: string }) {
    super({ filePath: './tslint.json', ...options });
  }

  get id(): string {
    return 'FN015003';
  }
}
