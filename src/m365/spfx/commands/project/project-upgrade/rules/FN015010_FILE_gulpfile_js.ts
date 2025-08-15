import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015010_FILE_gulpfile_js extends FileAddRemoveRule {
  constructor(add: boolean) {
    super('./gulpfile.js', add);
  }

  get id(): string {
    return 'FN015010';
  }
}
