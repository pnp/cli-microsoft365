import { FileAddRemoveRule } from "./FileAddRemoveRule";

export class FN015001_FILE_typings_tsd_d_ts extends FileAddRemoveRule {
  constructor(add: boolean) {
    super('./typings/tsd.d.ts', add);
  }

  get id(): string {
    return 'FN015001';
  }
}
