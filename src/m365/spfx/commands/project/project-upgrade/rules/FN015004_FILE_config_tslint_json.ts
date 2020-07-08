import { FileAddRemoveRule } from "./FileAddRemoveRule";

export class FN015004_FILE_config_tslint_json extends FileAddRemoveRule {
  constructor(add: boolean) {
    super('./config/tslint.json', add);
  }

  get id(): string {
    return 'FN015004';
  }

  get supersedes(): string [] {
    return ['FN008001', 'FN008002', 'FN008003'];
  }
}
