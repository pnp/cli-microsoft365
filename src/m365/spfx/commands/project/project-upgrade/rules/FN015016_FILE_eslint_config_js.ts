import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015016_FILE_eslint_config_js extends FileAddRemoveRule {
  constructor(add: boolean, contents: string) {
    super('./eslint.config.js', add, contents);
  }

  get id(): string {
    return 'FN015016';
  }
}
