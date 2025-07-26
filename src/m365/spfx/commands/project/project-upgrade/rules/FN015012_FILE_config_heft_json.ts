import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015012_FILE_config_heft_json extends FileAddRemoveRule {
  constructor(add: boolean, contents: string | undefined) {
    super('./config/heft.json', add, contents);
  }

  get id(): string {
    return 'FN015012';
  }
}
