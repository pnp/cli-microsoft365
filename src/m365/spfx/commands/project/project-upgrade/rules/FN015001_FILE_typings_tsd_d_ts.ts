import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015001_FILE_typings_tsd_d_ts extends FileAddRemoveRule {
  constructor(options: { add: boolean }) {
    super({
      filePath: './typings/tsd.d.ts',
      ...options
    });
  }

  get id(): string {
    return 'FN015001';
  }
}
