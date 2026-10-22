import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015010_FILE_gulpfile_js extends FileAddRemoveRule {
  constructor(options: { add: boolean }) {
    super({ filePath: './gulpfile.js', ...options });
  }

  get id(): string {
    return 'FN015010';
  }
}
