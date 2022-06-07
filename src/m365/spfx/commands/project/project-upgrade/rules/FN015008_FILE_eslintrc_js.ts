import { FileAddRemoveRule } from "./FileAddRemoveRule";

export class FN015008_FILE_eslintrc_js extends FileAddRemoveRule {
  constructor(add: boolean, contents: string) {
    super('./.eslintrc.js', add, contents);
  }

  get id(): string {
    return 'FN015008';
  }
}
