import { FileAddRemoveRule } from "./FileAddRemoveRule";

export class FN015002_FILE_typings__ms_odsp_d_ts extends FileAddRemoveRule {
  constructor(add: boolean) {
    super('./typings/@ms/odsp.d.ts', add);
  }

  get id(): string {
    return 'FN015002';
  }
}
