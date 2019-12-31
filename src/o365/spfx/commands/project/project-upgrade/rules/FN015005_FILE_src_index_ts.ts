import { FileAddRemoveRule } from "./FileAddRemoveRule";

export class FN015005_FILE_src_index_ts extends FileAddRemoveRule {
  constructor(add: boolean, contents: string) {
    super('./src/index.ts', add, contents);
  }

  get id(): string {
    return 'FN015005';
  }
}
