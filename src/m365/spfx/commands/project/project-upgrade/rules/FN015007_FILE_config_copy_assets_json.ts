import { FileAddRemoveRule } from "./FileAddRemoveRule";

export class FN015007_FILE_config_copy_assets_json extends FileAddRemoveRule {
  constructor(add: boolean) {
    super('./config/copy-assets.json', add);
  }

  get id(): string {
    return 'FN015007';
  }

  get supersedes(): string [] {
    return ['FN004001', 'FN004002'];
  }
}
