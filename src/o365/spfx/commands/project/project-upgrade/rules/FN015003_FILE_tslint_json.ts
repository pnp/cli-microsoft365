import { FileAddRemoveRule } from "./FileAddRemoveRule";

export class FN015003_FILE_tslint_json extends FileAddRemoveRule {
  constructor(add: boolean, contents: string) {
    /* istanbul ignore next */
    super('./tslint.json', add, contents);
  }

  get id(): string {
    return 'FN015003';
  }
}
