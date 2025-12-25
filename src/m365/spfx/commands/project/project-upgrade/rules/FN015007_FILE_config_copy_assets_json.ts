import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015007_FILE_config_copy_assets_json extends FileAddRemoveRule {
  constructor(options: { add: boolean }) {
    super({
      filePath: './config/copy-assets.json',
      add: options.add
    });
  }

  get id(): string {
    return 'FN015007';
  }

  get supersedes(): string[] {
    return ['FN004001', 'FN004002'];
  }
}
