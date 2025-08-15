import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015014_FILE_config_rig_json extends FileAddRemoveRule {
  constructor(add: boolean, contents: string | undefined) {
    super('./config/rig.json', add, contents);
  }

  get id(): string {
    return 'FN015014';
  }
}
