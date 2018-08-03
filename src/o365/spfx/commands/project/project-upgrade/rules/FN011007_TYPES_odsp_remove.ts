import { FileRemoveRule } from "./FileRemoveRule";

export class FN011007_TYPES_odsp_remove extends FileRemoveRule {
  public constructor() {
    super('/typings/@ms/odsp.d.ts');
  }
  get id(): string { return 'FN011007'; }
}
