import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

export class FN015002_FILE_typings__ms_odsp_d_ts extends FileAddRemoveRule {
  constructor(options: { add: boolean }) {
    super({ filePath: './typings/@ms/odsp.d.ts', ...options });
  }

  get id(): string {
    return 'FN015002';
  }
}
