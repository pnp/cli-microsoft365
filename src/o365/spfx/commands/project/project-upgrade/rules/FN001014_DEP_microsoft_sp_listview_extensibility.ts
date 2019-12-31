import { DependencyRule } from "./DependencyRule";

export class FN001014_DEP_microsoft_sp_listview_extensibility extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-listview-extensibility', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001014';
  }
}