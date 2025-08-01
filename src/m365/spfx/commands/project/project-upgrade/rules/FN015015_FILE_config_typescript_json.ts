import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015015_FILE_config_typescript_json extends FileAddRemoveRule {
  constructor(add: boolean, contents: string | undefined) {
    super('./config/typescript.json', add, contents);
  }

  get id(): string {
    return 'FN015015';
  }
}
