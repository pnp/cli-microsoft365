import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015011_FILE_tsconfig_json extends FileAddRemoveRule {
  constructor(add: boolean, contents: string | undefined) {
    super('./tsconfig.json', add, contents);
  }

  get id(): string {
    return 'FN015011';
  }
}
