import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015014_FILE_config_rig_json extends FileAddRemoveRule {
  constructor(options: { add: boolean; contents?: string }) {
    super({
      filePath: './config/rig.json',
      ...options
    });
  }

  get id(): string {
    return 'FN015014';
  }
}
