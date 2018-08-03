import { FileRemoveRule } from "./FileRemoveRule";

export class FN011010_TYPES_tsd_remove extends FileRemoveRule {
  public constructor() {
    super('/typings/tsd.d.ts');
  }
  get id(): string { return 'FN011010'; }
}
