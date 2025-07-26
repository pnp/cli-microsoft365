import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015013_FILE_config_jest_config_json extends FileAddRemoveRule {
  constructor(add: boolean, contents: string | undefined) {
    super('./config/jest.config.json', add, contents);
  }

  get id(): string {
    return 'FN015013';
  }
}
