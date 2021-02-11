import { FileAddRemoveRule } from "./FileAddRemoveRule";

export class FN015006_FILE_editorconfig extends FileAddRemoveRule {
  constructor(add: boolean) {
    super('.editorconfig', add);
  }

  get id(): string {
    return 'FN015006';
  }
}
