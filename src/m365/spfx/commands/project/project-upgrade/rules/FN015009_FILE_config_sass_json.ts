import { FileAddRemoveRule } from "./FileAddRemoveRule";

export class FN015009_FILE_config_sass_json extends FileAddRemoveRule {
  constructor(add: boolean, contents: string) {
    super('./config/sass.json', add, contents);
  }

  get id(): string {
    return 'FN015009';
  }
}
