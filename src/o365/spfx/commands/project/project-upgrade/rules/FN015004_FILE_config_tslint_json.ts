import { FileAddRemoveRule } from "./FileAddRemoveRule";

export class FN015004_FILE_config_tslint_json extends FileAddRemoveRule {
  constructor(add: boolean) {
    /* istanbul ignore next */
    super('./config/tslint.json', add);
  }

  get id(): string {
    return 'FN015004';
  }
}
