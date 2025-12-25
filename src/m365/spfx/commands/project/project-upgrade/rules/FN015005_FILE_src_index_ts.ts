import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015005_FILE_src_index_ts extends FileAddRemoveRule {
  constructor(options: { add: boolean; contents?: string }) {
    super({
      filePath: './src/index.ts',
      ...options
    });
  }

  get id(): string {
    return 'FN015005';
  }
}
